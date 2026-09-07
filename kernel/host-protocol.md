# Kernel host protocol 1

One public invocation function is used by both hosts: `kernelInvoke<T>(operation: string, params: unknown): T`. It is synchronous after `await initializeKernel()` and throws the typed kernel error on failure. There is no fallback evaluator. The WASM ABI exports `kernel_alloc`, `kernel_free`, `kernel_invoke`, `kernel_result_ptr`, `kernel_result_len`; invoke accepts UTF-8 JSON `{protocolVersion:1,requestId,operation,params}` and writes the same envelope used by native.

Native control messages are framed by a 4-byte big-endian length, maximum 16 MiB. Response is `{protocolVersion:1,requestId,ok,result?,error?:{code,message,object,recovery}}`. Operation names below are canonical, without aliases.

* `init`: `{}` -> `{protocolVersion:1,manifestVersion:11,operations:[...]}`.
* `open`: `{manifest:WorkbookManifest,pages?:PagePayload[]}` loads a revision-pinned replica; all pages listed in the manifest but absent from this request are unavailable, not blank.
* `create`: `{unitId,name,sheets?:[SheetManifest]}` -> WorkbookManifest at revision 0; no persistent state until Java commits it. Omitted sheets creates the native-owned `sheet-1` / `Sheet1` with full Excel row/column capacity and no allocated pages. Explicit null, empty or invalid sheets are rejected; existing identities are never replaced.
* `manifest`: `{unitId,revision}` -> WorkbookManifest.
* `copy`: `{sourceUnitId,sourceRevision,targetUnitId,name}` -> `{manifest}`. Creates a distinct revision-zero identity with unchanged page content hashes, sheet identities and authored metadata. The target's pages remain unavailable until Java copies verified durable content and loads requested pages. Existing targets and stale sources are rejected. Java copies native artifact provenance by exporting copied canonical pages into the preserved source package at the new identity boundary in the same transaction.
* `cell.get`: `{unitId,revision,address:{sheetId,row,column}}` -> `{revision,cell:null|Cell}`.
* `range.get`: `{unitId,revision,range:RangeRef}` -> `{revision,cells:[{address,cell}]}`; bounded range response, use pages for bulk data.
* `page.get`: `{unitId,revision,sheetId,pageRow,pageColumn}` -> PagePayload.
* `page.load`: `{unitId,revision,page:PagePayload}` loads a proven manifest page into a replica. It cannot introduce a new page identity or mutate the committed manifest.
* `command`: `{unitId,baseRevision,operationId,commandId,params}` -> ChangeSet. Identity/ACL is supplied and checked by Java; client actor claims never grant authority.
* `close`: `{unitId}` -> `{closed:true}`.
* `restore`: `{unitId,baseRevision,operationId,accessRole:"owner",targetManifest}` -> ChangeSet at baseRevision+1. Target must belong to the same workbook and cannot be from the future. Historical immutable page references may appear without inline payloads; Java verifies their durable content before committing. Clients hydrate changed references through the revision-pinned page endpoint.
* `formula.evaluate`: `{unitId,revision,address,formula}` -> `{revision,value}`.
* `formula.recalculate`: `{unitId,revision}` -> `{revision,values:[{address,value}]}`.
* `analytics.execute`: `{unitId,revision,request:AnalyticsRequest}` -> AnalyticsResponse.
* `geometry.computePaneMap`: GeometryRequest -> GeometryResponse.
* `geometry.hitTest`: `{request:GeometryRequest,point:Point}` -> HitTestResponse.
* `geometry.cellRect`: `{request:GeometryRequest,address:CellAddress}` -> Rect.
* `geometry.headerRect`: `{request:GeometryRequest,axis,index?}` -> HeaderRectResponse.
* `document.import` and `document.export`: native host only. Large artifacts use an explicitly configured task-directory file handle, not inline base64 files. Root/Java/native document owners coordinate this contract before activation.

WorkbookManifest = `{schema:"WorkbookManifest",version:11,unitId,name,revision,sheets:[{sheetId,name,rowCount,columnCount,metadata}],pages:[PageDescriptor],metadata}`. PageDescriptor = `{sheetId,pageRow,pageColumn,revision,checksum,byteLength}`. PagePayload extends PageDescriptor with `payloadBase64` (bounded kernel page encoding). Cell authored metadata remains flattened, matching Rust core Cell. No dense `cells` collection is present on the manifest.

ChangeSet = `{operationId,baseRevision,revision,manifest,pages:[PagePayload],removedPages:[PageKey],affectedRanges:[RangeRef]}`. Java commits this with operation/checkpoint/outbox, then broadcasts. Native staged state must not be treated as committed if the database transaction fails; reopen the last committed manifest before the next operation.

The migration executable is the only reader of v10 full snapshots. Runtime host accepts manifest v11 only. Columnar page records are data transfer, not a parallel canonical model.
