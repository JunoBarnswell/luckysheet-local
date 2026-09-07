use crate::{KernelHost, MAX_FRAME_BYTES};
use std::cell::RefCell;

thread_local! {
    static HOST: RefCell<KernelHost> = RefCell::new(KernelHost::default());
    static RESULT: RefCell<Vec<u8>> = const { RefCell::new(Vec::new()) };
}

#[unsafe(no_mangle)]
pub extern "C" fn kernel_alloc(length: u32) -> u32 {
    if length == 0 || length as usize > MAX_FRAME_BYTES {
        return 0;
    }
    let bytes = vec![0_u8; length as usize].into_boxed_slice();
    Box::into_raw(bytes) as *mut u8 as u32
}

/// Only buffers returned by kernel_alloc may be released, once, with the original length.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn kernel_free(pointer: u32, length: u32) {
    if pointer == 0 || length == 0 || length as usize > MAX_FRAME_BYTES {
        return;
    }
    let slice = std::ptr::slice_from_raw_parts_mut(pointer as *mut u8, length as usize);
    unsafe {
        drop(Box::from_raw(slice));
    }
}

#[unsafe(no_mangle)]
pub unsafe extern "C" fn kernel_invoke(pointer: u32, length: u32) -> u32 {
    if pointer == 0 || length == 0 || length as usize > MAX_FRAME_BYTES {
        return 1;
    }
    let request = unsafe { std::slice::from_raw_parts(pointer as *const u8, length as usize) };
    let response = HOST.with(|host| host.borrow_mut().invoke(request));
    RESULT.with(|result| *result.borrow_mut() = response);
    0
}

#[unsafe(no_mangle)]
pub extern "C" fn kernel_result_ptr() -> u32 {
    RESULT.with(|result| result.borrow().as_ptr() as u32)
}
#[unsafe(no_mangle)]
pub extern "C" fn kernel_result_len() -> u32 {
    RESULT.with(|result| result.borrow().len() as u32)
}
