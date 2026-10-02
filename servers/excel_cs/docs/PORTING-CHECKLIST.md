# Checklist tiến độ port Excel MCP sang C#

**Ảnh chụp tiến độ:** cập nhật 2026-10-02 trên `feat/excel-cs-port` (bao gồm `expect.value`/`expect.formula` và kịch bản parity hẹp). Đây là checklist thực thi; hợp đồng đầy đủ nằm trong [01](01-architecture.md), [03](03-tools.md), [04](04-correctness.md), [05](05-legacy-mapping.md) và [06](06-testing-and-roadmap.md). Dấu `[x]` chỉ khẳng định **phạm vi ghi ngay trong mục đó**, không có nghĩa toàn bộ phase đã hoàn tất.

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
- [x] `excel_apply`: `set_values` hình chữ nhật, `set_value` broadcast, `set_formula` thông thường (clear/keep/explicit typed cache), `fill` hằng số hoặc chuỗi thập phân có giới hạn, `clear` **chỉ values** (có thể xoá cell); giới hạn tối đa 500 cell/batch.
- [x] G6 từ chối source có chữ ký và bảo vệ các advanced parts trong phạm vi hỗ trợ; G7 trên staging hỗ trợ `equals.value`, `equals.formula`, và `unchanged: true` cho cell (kể cả shared-string markup/phonetic hoặc cell vắng mặt).
- [ ] Mở rộng G6 cho thay đổi advanced parts có chủ đích và chính sách invalidation chữ ký; G7 cho các facet `display`/`rich`/`style`/`except`, với phép đọc độc lập và test âm tương ứng.
- [x] `expect.value`, `expect.formula` (riêng hoặc cùng nhau) cho `set_value`/`set_formula`/`clear` **một cell** kiểm trước batch trên revision hiện tại; mismatch trả `PRECONDITION_FAILED` có index của **op gốc** (kể cả sau range expansion), expected/actual, không đổi revision.
- [ ] Hoàn tất rich-text/phonetic và style ops, các facet `expect` còn lại (kể cả bulk/hash), `dry_run`, readback/diff đầy đủ; không coi các op `set_value` hiện có là parity của toàn bộ 03 §4.
- [ ] `excel_create`, các chế độ save `overwrite`/`save_as`, persistent ledger/redo (nếu quyết định hỗ trợ) và các option session còn thiếu; hiện **chỉ** cho copy-save.
- [ ] Chạy ma trận op × fixture, parity Python/C# có danh sách sai khác được duyệt, và agent-eval subset đạt **0 silent failure** theo [06 §2.7](06-testing-and-roadmap.md#27-agent-level-evals). Đã có [kịch bản parity hẹp](../tools/PARITY.md) qua 6 fixture cho scalar edits; 1 Python output tương đương, 5 divergence được kiểm đúng theo fixture; chưa phải ma trận đầy đủ.

## 3. Phần còn lại của mục tiêu port

- [ ] **P0/P1 đầy đủ:** chốt các spike S1–S9 chưa đạt exit criteria; hoàn tất `excel_peek`/`find`/`inspect`/`verify` độc lập, các view/read semantics còn thiếu; kiểm trên corruption/equivalence corpus và đối chiếu Python. Những phần đang phục vụ P2a không tự động hoàn thành P1.
- [ ] **P3:** thêm row/column, merge, sheet ops, transform-aware verification và formula shifter; differential fuzzing qua ngưỡng đã chốt.
- [ ] **P4:** tables, conditional formatting, data validation, names, hyperlinks, comments, printing, views, protection, workbook/doc properties.
- [ ] **P5:** drawings, package/expert ops và `excel_export`, đối chiếu kịch bản EX-01.
- [ ] **P6:** Excel/LibreOffice oracle (render/recalc/open-check), benchmark, agent eval đầy đủ, self-contained binaries và clean-machine smoke.
- [ ] **P7:** hoàn tất hoặc quyết định bỏ rõ ràng từng mục trong [mapping 112 tool Python](05-legacy-mapping.md), không còn lỗi nghiêm trọng, rồi **mới** chuyển launcher `excel` sang C#; giữ Python dưới tên tương thích theo lộ trình.

## 4. Bằng chứng và cách cập nhật checklist

- [x] Ở mốc `9b385dc`, local Windows và Docker Linux chạy **315/315** test. Lát cắt `expect.value` qua **330/330**, sau khi thêm `expect.formula` qua **339/339** test trên cả Windows và Docker Linux (có nguồn fixture local). Sửa index sau range expansion đạt **345/345** test trên Windows và Docker Linux, MCP stdio smoke; parity hẹp chạy trên Windows cho 6 fixture (1 Python tương đương, 5 khác biệt đã ghi). Schema-order check đã qua ở lát cắt trước. Đây chỉ là bằng chứng cho các lát cắt hiện có, không chứng minh P2b/P3–P7.
- [ ] Chạy CI trên **cả bốn runner** cho commit mới (workflow đã cấu hình, nhưng các commit local này chưa push); macOS local được hoãn theo yêu cầu, không đánh dấu đã kiểm chứng.
- [ ] S3: đo **model thực nhận** `structuredContent` hay `TextContent` trên từng client đích; smoke SDK chỉ chứng minh giao thức tới client, không chứng minh forwarding vào model.
- [ ] Ma trận `set_value` với corpus cục bộ `D:\data-test\excel-preservation-fixtures\sources` và bộ tổng hợp: lưu báo cáo pass/fail theo từng fixture/gate trước khi đóng tiêu chí release; ghi rõ 01 lỗi schema baseline, 02/05 Excel từ chối ngay file gốc ([S2b](../spikes/S2b/REPORT.md)).
- [ ] Thử release trên Windows/Linux và sau đó macOS/clean-machine trước khi tuyên bố khả năng thay Python trên mọi OS.

Chạy lại từ root repo (test local fixture là tuỳ chọn, không được thêm các file này vào commit):

```powershell
$env:DOCLOUPE_P2A_LOCAL_FIXTURES = 'D:\data-test\excel-preservation-fixtures\sources'
dotnet test servers/excel_cs/DocLoupe.Excel.slnx --configuration Release
dotnet run --project servers/excel_cs/tools/SchemaOrder/SchemaOrder.csproj --configuration Release
dotnet run --project servers/excel_cs/spikes/S3/S3.csproj --configuration Release -- servers/excel_cs/src/DocLoupe.Excel.Server/bin/Release/net10.0/DocLoupe.Excel.Server.dll
```

**Thứ tự công việc tiếp theo:** (1) mở rộng [kịch bản parity](../tools/PARITY.md) cho formula/clear, rich/style và corpus độc lập; (2) hoàn tất reader và các op rich/style kèm G4–G7; (3) đóng tiêu chí P0/P1/P2b chưa đạt; (4) P3 → P7 theo [roadmap](06-testing-and-roadmap.md#4-roadmap). Mỗi mục chỉ chuyển thành `[x]` khi có test/báo cáo dẫn chứng và commit tương ứng; không push nếu chưa được yêu cầu.
