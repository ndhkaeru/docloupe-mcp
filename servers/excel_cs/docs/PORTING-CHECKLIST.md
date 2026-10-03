# Checklist tiến độ port Excel MCP sang C#

**Ảnh chụp tiến độ:** cập nhật 2026-10-03 trên `feat/excel-cs-port` (bao gồm `expect.value`/`expect.formula` và kịch bản parity hẹp). Đây là checklist thực thi; hợp đồng đầy đủ nằm trong [01](01-architecture.md), [03](03-tools.md), [04](04-correctness.md), [05](05-legacy-mapping.md) và [06](06-testing-and-roadmap.md). Dấu `[x]` chỉ khẳng định **phạm vi ghi ngay trong mục đó**, không có nghĩa toàn bộ phase đã hoàn tất.

**Đích cuối:** server C# thay thế được `servers/excel` (Python) mà không mất chỉnh sửa hoặc phá hỏng nội dung OOXML; trước khi đổi launcher cần đạt tiêu chí P7 trong [06 §4](06-testing-and-roadmap.md#4-roadmap). **Hiện tại: P2a chạy được; P2b đang làm; chưa thể thay server Python.**

| Mốc | Trạng thái | Điều kiện còn thiếu để đóng mốc |
|---|---|---|
| P0 | Một phần | Các spike S1–S9 còn mở |
| P1 | Một phần | Reader/verifier tổng quát và parity corpus |
| P2a | Lát cắt hoạt động | Kiểm nốt điều kiện nghiệm thu với corpus local/CI thực tế |
| P2b | Đang thực hiện | Rich/style, full G6–G7, create/save modes và agent eval |
| P3 | Chưa bắt đầu | Thao tác cấu trúc và công thức liên quan |
| P4 | Chưa bắt đầu | Tính năng workbook/worksheet nâng cao |
| P5 | Chưa bắt đầu | Drawing và package ops |
| P6 | Chưa bắt đầu | Oracle, hiệu năng, phát hành và clean-machine |
| P7 | Chưa bắt đầu | Parity đầy đủ và chuyển launcher |

## 1. Nền tảng và P2a — đã triển khai trong phạm vi hẹp

- [x] Solution .NET 10 gồm `Model`, `Package`, `Engine`, `Verify`, `Schema`, `Server` và test; test kiến trúc giữ `Verify` độc lập với `Engine`/Open XML SDK.
- [x] `PackageStore` tìm workbook qua OPC relationships, xử lý tên part không phân biệt hoa thường và percent-encoding; giữ nguyên nội dung entry không sửa.
- [x] Bộ ghi `set_value`: phân tích A1, định vị theo namespace và `r`, chèn row/cell theo thứ tự schema, ghép nhiều byte-span từ offset giảm dần, hỗ trợ số/chuỗi/shared string/inline string/công thức trong phạm vi P2a.
- [x] Chuỗi mới dùng shared strings; cell inline hiện hữu giữ inline. Đổi giá trị bật `fullCalcOnLoad`; chỉ xoá `calcChain` khi ghi đè/xoá công thức hiện hữu.
- [x] `excel_open` → `excel_read` (cell đích) → `excel_apply` → `excel_save` (**copy ra đường dẫn mới**) → `excel_close` chạy qua MCP stdio; save ghi staging, chạy gate rồi mới chuyển ra đích.
- [x] G1 (liên kết/`r:id`), G2 (validate tách rời và báo lỗi mới/lỗi bị che), G3 (namespace/markup), G4 (ý định và đọc lại, kể cả ca mất edit), G5 (byte ngoài span và cell đụng tới) có test phá hỏng cố ý trong phạm vi writer hiện có; **không đánh dấu là verifier tổng quát**.
- [x] Sáu fixture tổng hợp sinh từ source trong repo: namespace mặc định/`x:`, `ns0:Types`, BOM/CRLF/standalone, shared strings, cache công thức, phonetic, `mc`/`x14ac`, OPC percent/case, workbook lồng. CI không phụ thuộc `D:\data-test`.
- [x] Workflow `excel-cs.yml` khai báo Windows x64, Linux x64, macOS x64 và macOS arm64, chạy test, schema-order check và MCP smoke.

## 2. P2b — đã có từng lát cắt, chưa xong phase

- [x] `excel_status` (revision, dirty, thay đổi source, busy, ledger 20 mục) và `excel_undo` bằng phát lại những revision cell-content còn giữ; copy-save **không** đánh dấu phiên là đã lưu.
- [x] `excel_read` một ô/vùng A1 hoặc danh sách target giới hạn tổng 500 ô, chỉ `cells` hiện hữu trên revision đang mở; chưa hỗ trợ `values`/`markdown`/`full`, hash hoặc phân trang. `skip_empty=false` trong view `cells` trả placeholder `blank` có địa chỉ theo thứ tự target; `skip_empty=true` vẫn là mặc định.
- [x] Reader P1 `WorkbookReader.Peek` định vị workbook từ `_rels/.rels`, hỗ trợ workbook lồng, đường dẫn percent-encoded và tên part khác chữ hoa/thường; `VerifyPartial` báo lỗi root/workbook relationship, worksheet part thiếu, ZIP hỏng, cùng target hỏng/ID trùng/source thiếu trên mọi `.rels` và `r:id`/`r:embed`/`r:link`/`r:pict` trong XML theo namespace cùng content type hiệu lực, loại workbook theo đuôi file và preflight ZIP theo 01 §7. Kiểm thêm tọa độ row/cell sai, trùng hoặc không khớp hàng trên **các worksheet có tham chiếu** (giữ `r` tùy chọn hợp lệ, suy ra vị trí từ row/cell trước đó để phát hiện trùng hoặc không khớp); G1 tổng quát vẫn mở. Preview giữ raw `<v>` và đọc giá trị shared string (bỏ phonetic); chỉ số shared string hỏng bị từ chối thay vì bị coi là ô rỗng. `VerifyPartial` kiểm chỉ số `t="s"` trên mọi sheet được tham chiếu (kể cả ngoài preview), đếm `<si>` từ part shared strings theo OPC mà không tải toàn bộ nội dung chuỗi; thiếu relationship hoặc vượt giới hạn đều báo lỗi. Read-only `Peek` suy ra vị trí row/cell khi `r` thiếu và giữ nguyên địa chỉ tường minh; writer vẫn từ chối phần tử không xác định được vị trí theo yêu cầu P2a. `excel_verify(after_path)` chỉ nối các kiểm tra này với MCP, trả `failed`/`isError` cho file hỏng, `unverified`/`partial` khi chưa thấy lỗi; **không** phải `excel_verify` hoàn chỉnh.
- [x] `excel_apply`: `set_values` hình chữ nhật, `set_value` broadcast, `set_formula` thông thường (clear/keep/explicit typed cache), `fill` hằng số hoặc chuỗi thập phân có giới hạn, `clear` **chỉ values** (có thể xoá cell); giới hạn tối đa 500 cell/batch.
- [x] G6 từ chối source có chữ ký và bảo vệ các advanced parts trong phạm vi hỗ trợ; G7 trên staging hỗ trợ `equals.value`, `equals.formula`, và `unchanged: true` cho cell (kể cả shared-string markup/phonetic hoặc cell vắng mặt).
- [ ] Mở rộng G6 cho thay đổi advanced parts có chủ đích và chính sách invalidation chữ ký; G7 cho các facet `display`/`rich`/`style`/`except`, với phép đọc độc lập và test âm tương ứng.
- [x] `expect.value`, `expect.formula` (riêng hoặc cùng nhau) cho `set_value`/`set_formula`/`clear` **một cell** kiểm trước batch trên revision hiện tại; mismatch trả `PRECONDITION_FAILED` có index của **op gốc** (kể cả sau range expansion), expected/actual, không đổi revision.
- [x] `excel_apply.dry_run` chạy cùng precondition và candidate writer nhưng không thay revision/ledger/nguồn; trả `intent`/`changed_parts` nguồn-tương-đối (gồm các edit trước đó) và revision trước/sau bằng nhau. Chưa trả diff/readback đầy đủ theo 03 §4.
- [ ] Hoàn tất rich-text/phonetic và style ops, các facet `expect` còn lại (kể cả bulk/hash), readback/diff đầy đủ; không coi các op `set_value` hiện có là parity của toàn bộ 03 §4.
- [ ] `excel_create`, các chế độ save `overwrite`/`save_as`, persistent ledger/redo (nếu quyết định hỗ trợ) và các option session còn thiếu; hiện **chỉ** cho copy-save.
- [ ] Chạy ma trận op × fixture, parity Python/C# có danh sách sai khác được duyệt, và agent-eval subset đạt **0 silent failure** theo [06 §2.7](06-testing-and-roadmap.md#27-agent-level-evals). Đã có [kịch bản parity hẹp](../tools/PARITY.md) qua 6 fixture cho scalar edits, formula/cache, value-only clear; 1 Python output tương đương, 5 divergence được kiểm đúng theo fixture trên Windows và Docker Linux; chưa phải ma trận đầy đủ.

## 3. Phần còn lại của mục tiêu port

- [ ] **P0/P1 đầy đủ:** chốt các spike S1–S9 chưa đạt exit criteria; `excel_peek` hiện trả info/summary tối giản với `used_range` dựa vào explicit cell (không tính row-only formatting) và preview Markdown đọc-only (tối đa 100 hàng × 20 cột/2.000 ô, báo `partial` và `unverified` cho features); chưa đủ hợp đồng `features` hoặc `excel_find`/`excel_inspect`/`excel_verify` đầy đủ (verify read-only có G2 tách rời cho workbook/worksheet/shared strings, báo gap cho XML root khác; compare read-only chỉ so byte decompressed từng part, chưa so facet/normalization; `assert` có thể bắt edit bị mất khi hai file giống nhau, nhưng không suy ra được ý định nếu thiếu assertion), các view/read semantics còn thiếu; cần kiểm corruption/equivalence corpus và đối chiếu Python. Những phần đang phục vụ P2a không tự động hoàn thành P1.
- [ ] **P3:** thêm row/column, merge, sheet ops, transform-aware verification và formula shifter; differential fuzzing qua ngưỡng đã chốt.
- [ ] **P4:** tables, conditional formatting, data validation, names, hyperlinks, comments, printing, views, protection, workbook/doc properties.
- [ ] **P5:** drawings, package/expert ops và `excel_export`, đối chiếu kịch bản EX-01.
- [ ] **P6:** Excel/LibreOffice oracle (render/recalc/open-check), benchmark, agent eval đầy đủ, self-contained binaries và clean-machine smoke.
- [ ] **P7:** hoàn tất hoặc quyết định bỏ rõ ràng từng mục trong [mapping 112 tool Python](05-legacy-mapping.md), không còn lỗi nghiêm trọng, rồi **mới** chuyển launcher `excel` sang C#; giữ Python dưới tên tương thích theo lộ trình.

## 4. Bằng chứng và cách cập nhật checklist

- [x] Ở mốc `9b385dc`, local Windows và Docker Linux chạy **315/315** test. Lát cắt `expect.value` qua **330/330**, sau khi thêm `expect.formula` qua **339/339** test trên cả Windows và Docker Linux (có nguồn fixture local). Sửa index sau range expansion đạt **345/345** test; đọc A1 range giới hạn đạt **358/358** test trên Windows và Docker Linux, MCP stdio smoke; reader OPC/worksheet-link, shared-string preview và `excel_peek` P1 đạt **388/388** test trên Windows và Docker Linux (sáu fixture tổng hợp); quét `used_range` theo cell thực, bỏ qua `<dimension>` cũ; MCP stdio smoke qua cả hai OS; `excel_verify` partial đọc-only (file tốt, root `.rels` thiếu, ZIP hỏng) đạt **389/389** test trên Windows (có corpus local) và Docker Linux, MCP stdio smoke qua cả hai OS; G1 partial kiểm `.rels` part phụ (target, ID trùng, source thiếu, external/internal) đạt **396/396** test trên Windows (có corpus local) và Docker Linux, MCP stdio smoke qua cả hai OS; `r:id` theo namespace trên XML part đạt **400/400** test trên Windows (có corpus local) và Docker Linux, MCP stdio smoke qua cả hai OS; `r:embed`/`r:link`/`r:pict` đạt **403/403** test trên Windows (có corpus local) và Docker Linux, MCP stdio smoke qua cả hai OS; content type hiệu lực/loại workbook đạt **415/415** test trên Windows (có corpus local) và Docker Linux, MCP stdio smoke qua cả hai OS; preflight ZIP theo 01 §7 đạt **418/418** test trên Windows (có corpus local) và Docker Linux, MCP stdio smoke qua cả hai OS; G2 read-only trên các root workbook/worksheet/sst và báo gap cho root chưa hỗ trợ; fixture tổng hợp đã sửa hai lỗi schema (sheetViews/sheetFormatPr), bộ **423/423** test đạt trên Windows và Docker Linux (gồm test âm cho workbook/worksheet/sst và unsupported root); compare part-level read-only đạt **426/426** test trên Windows và Docker Linux; so schema delta với lỗi baseline (kể cả masked baseline) đạt **427/427** test; assertion read-only bắt edit bị mất ngay cả khi hai file giống nhau đạt **428/428** test trên Windows và Docker Linux, MCP stdio smoke qua cả hai OS; parity hẹp chạy trên Windows và Docker Linux cho 6 fixture (1 Python tương đương, 5 khác biệt đã ghi). Bản `excel_peek` này vẫn chỉ là lát cắt read-only; `used_range` có basis `explicit_cells` (không gồm row-only formatting), còn `features` được báo `unverified`. Đây chỉ là bằng chứng cho các lát cắt hiện có, không chứng minh P2b/P3–P7.
- [x] 2026-10-03: kiểm tọa độ trên mọi worksheet được tham chiếu và từ chối chỉ số shared-string sai trong readback/G7; Windows và Docker Linux đạt **439/439** test (Windows có corpus local), schema-order và MCP stdio smoke qua cả hai OS. Kiểm này chỉ bao phủ tọa độ tường minh và các cell G7 đọc, không xác nhận toàn bộ P1/P2b.
- [x] 2026-10-03: `VerifyPartial` kiểm chỉ số `t="s"` cả trên sheet ngoài preview, báo lỗi nếu thiếu shared-strings relationship. Windows và Docker Linux đạt **445/445** test với corpus local, schema-order và MCP stdio smoke qua cả hai OS. Chưa xác nhận nội dung/định dạng mọi `<si>` bằng kiểm tra này.
- [x] 2026-10-03: `WorkbookReader.Peek` suy ra row/cell khi `r` tùy chọn vắng mặt theo thứ tự hiện hữu, kể cả cột trống giữa hai cell tường minh; lỗi nếu suy ra vượt giới hạn. Windows và Docker Linux đạt **448/448** test với corpus local, schema-order và MCP stdio smoke qua cả hai OS. Writer vẫn fail-closed trên cell thiếu `r`.
- [x] 2026-10-03: `VerifyPartial` kiểm tọa độ suy ra khi thiếu `r`, gồm hàng/ô trùng, hàng không khớp, vượt giới hạn và shared string hỏng. Windows và Docker Linux đạt **456/456** test với corpus local, schema-order và MCP stdio smoke; parity Python/C# vẫn đạt đúng 6 kịch bản hiện hữu trên Windows (1 tương đương, 5 divergence đã ghi).
- [x] 2026-10-03: `excel_peek` dùng chung preflight ZIP 01 §7 trước khi giải nén dữ liệu worksheet; test từ chối ZIP quá nhiều entry và entry nén quá mức. **456/456** test Windows/Docker Linux với corpus local, schema-order và MCP stdio smoke qua hai OS.
- [x] 2026-10-03: `P2aGates.ReadCells` đọc tọa độ suy ra ngay trong `sheetData` cho `excel_read`/G7, không coi cell thiếu `r` là không tồn tại. Windows và Docker Linux đạt **457/457** test với corpus local, schema-order và MCP stdio smoke; writer vẫn fail-closed với cell không có `r`.
- [x] 2026-10-03: `excel_read` `cells` hỗ trợ `skip_empty=false` (placeholder `blank` theo thứ tự, kể cả target trùng); mặc định vẫn bỏ ô rỗng. Windows/Docker Linux đạt **458/458** test với corpus local, schema-order và MCP stdio smoke.
- [x] 2026-10-03: `PackageStore` giải phóng ZIP khi constructor thất bại (root thiếu, part trùng, relationship lỗi); test mở file độc quyền sau lỗi. Windows/Docker Linux đạt **461/461** test với corpus local, schema-order và MCP stdio smoke.
- [x] 2026-10-03: `excel_apply.dry_run` có 7 test mới trên fixture tổng hợp, so kế hoạch với apply thật và kiểm session/source không đổi, cùng test lỗi revision/precondition/writer; **468/468** test trên Windows và Docker Linux (gồm corpus local), schema-order qua hai OS; smoke MCP có ca `dry_run`. Chưa có diff/readback theo hợp đồng đầy đủ.
- [ ] Chạy CI trên **cả bốn runner** cho commit mới (workflow đã cấu hình, nhưng các commit local này chưa push); macOS local được hoãn theo yêu cầu, không đánh dấu đã kiểm chứng.
- [ ] S3: đo **model thực nhận** `structuredContent` hay `TextContent` trên từng client đích; smoke SDK chỉ chứng minh giao thức tới client, không chứng minh forwarding vào model.
- [ ] Ma trận `set_value` với corpus cục bộ `D:\data-test\excel-preservation-fixtures\sources` và bộ tổng hợp: **8/8 test `LocalFixtureTests` qua trên Windows** (gồm 8 nguồn, set/clear/error/fill, chữ ký và reader `used_range`); vẫn thiếu báo cáo pass/fail theo từng fixture/gate trước khi đóng tiêu chí release. Ghi rõ 01 lỗi schema baseline, 02/05 Excel từ chối ngay file gốc ([S2b](../spikes/S2b/REPORT.md)).
- [ ] Thử release trên Windows/Linux và sau đó macOS/clean-machine trước khi tuyên bố khả năng thay Python trên mọi OS.

Chạy lại từ root repo (test local fixture là tuỳ chọn, không được thêm các file này vào commit):

```powershell
$env:DOCLOUPE_P2A_LOCAL_FIXTURES = 'D:\data-test\excel-preservation-fixtures\sources'
dotnet test servers/excel_cs/DocLoupe.Excel.slnx --configuration Release
dotnet run --project servers/excel_cs/tools/SchemaOrder/SchemaOrder.csproj --configuration Release
dotnet run --project servers/excel_cs/spikes/S3/S3.csproj --configuration Release -- servers/excel_cs/src/DocLoupe.Excel.Server/bin/Release/net10.0/DocLoupe.Excel.Server.dll
```

**Thứ tự công việc tiếp theo:** (1) mở rộng [kịch bản parity](../tools/PARITY.md) cho rich/style và corpus độc lập; (2) hoàn tất reader và các op rich/style kèm G4–G7; (3) đóng tiêu chí P0/P1/P2b chưa đạt; (4) P3 → P7 theo [roadmap](06-testing-and-roadmap.md#4-roadmap). Mỗi mục chỉ chuyển thành `[x]` khi có test/báo cáo dẫn chứng và commit tương ứng; không push nếu chưa được yêu cầu.
