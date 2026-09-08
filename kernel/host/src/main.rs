#[cfg(not(target_arch = "wasm32"))]
fn main() {
    use kernel_host::{KernelHost, MAX_FRAME_BYTES};
    use serde_json::{Value, json};
    use std::collections::BTreeMap;
    use std::io::{self, Read, Write};
    use std::sync::{Arc, Mutex, atomic::{AtomicBool, Ordering}, mpsc};

    fn write_frame(output: &Arc<Mutex<io::Stdout>>, bytes: &[u8]) {
        let mut writer = output.lock().expect("stdout lock poisoned");
        if writer.write_all(&(bytes.len() as u32).to_be_bytes())
            .and_then(|_| writer.write_all(bytes)).and_then(|_| writer.flush()).is_err() {
            std::process::exit(1);
        }
    }
    struct Job { id: String, bytes: Vec<u8>, cancel: Arc<AtomicBool> }
    let output = Arc::new(Mutex::new(io::stdout()));
    let cancellations: Arc<Mutex<BTreeMap<String, Arc<AtomicBool>>>> = Arc::new(Mutex::new(BTreeMap::new()));
    let (sender, receiver) = mpsc::sync_channel::<Job>(16);
    let worker_output = output.clone();
    let worker_cancellations = cancellations.clone();
    let worker = std::thread::spawn(move || {
        let mut host = KernelHost::default();
        while let Ok(job) = receiver.recv() {
            let response = host.invoke_cancellable(&job.bytes, job.cancel);
            write_frame(&worker_output, &response);
            worker_cancellations.lock().expect("cancellation lock poisoned").remove(&job.id);
        }
    });
    let mut input = io::stdin().lock();
    loop {
        let mut header = [0_u8; 4];
        match input.read_exact(&mut header) {
            Ok(()) => {},
            Err(error) if error.kind() == io::ErrorKind::UnexpectedEof => break,
            Err(error) => { eprintln!("KERNEL_INPUT_FAILED: {error}"); std::process::exit(1); }
        }
        let length = u32::from_be_bytes(header) as usize;
        if length == 0 || length > MAX_FRAME_BYTES {
            eprintln!("KERNEL_PAYLOAD_TOO_LARGE: frame length {length}");
            std::process::exit(2);
        }
        let mut bytes = vec![0; length];
        if let Err(error) = input.read_exact(&mut bytes) {
            eprintln!("KERNEL_FRAME_TRUNCATED: {error}"); std::process::exit(2);
        }
        let envelope: Option<Value> = serde_json::from_slice(&bytes).ok();
        let request_id = envelope.as_ref().and_then(|v|v["requestId"].as_str()).unwrap_or("").to_owned();
        if let Some(request) = &envelope {
            if request["protocolVersion"].as_u64() == Some(1) && request["operation"] == "task.cancel" {
                let task = request["params"]["taskRequestId"].as_str().unwrap_or("");
                let found = cancellations.lock().expect("cancellation lock poisoned").get(task).cloned();
                if let Some(flag) = &found { flag.store(true, Ordering::Release); }
                // notifyOnly is for the serialized Java client: the original request
                // supplies the terminal cancellation response, so no unrelated reply
                // can be mistaken for that request's result.
                if request["params"]["notifyOnly"] != true {
                    let response = json!({"protocolVersion":1,"requestId":request_id,"ok":true,"result":{"accepted":found.is_some(),"taskRequestId":task}});
                    write_frame(&output, &serde_json::to_vec(&response).unwrap());
                }
                continue;
            }
        }
        let cancel = Arc::new(AtomicBool::new(false));
        {
            let mut registry = cancellations.lock().expect("cancellation lock poisoned");
            if registry.contains_key(&request_id) {
                let error = kernel_core::KernelError::new("KERNEL_REQUEST_DUPLICATE", "A request with this identity is already active");
                write_frame(&output, &serde_json::to_vec(&json!({"protocolVersion":1,"requestId":request_id,"ok":false,"error":error})).unwrap());
                continue;
            }
            registry.insert(request_id.clone(),cancel.clone());
        }
        if sender.send(Job{id:request_id,bytes,cancel}).is_err() { std::process::exit(1); }
    }
    for flag in cancellations.lock().expect("cancellation lock poisoned").values() { flag.store(true,Ordering::Release); }
    drop(sender);
    if worker.join().is_err() { eprintln!("KERNEL_EXECUTION_FAILED: worker panicked"); std::process::exit(1); }
}
#[cfg(target_arch = "wasm32")]
fn main() {}
