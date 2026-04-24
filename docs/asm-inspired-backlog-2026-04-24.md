# ASM-Inspired Backlog

Ngày rà soát: `2026-04-24`

Phạm vi:
- Backlog này lấy ý tưởng từ các merged PR gần đây của `luongnv89/asm`, nhưng chỉ giữ những gì áp dụng trực tiếp cho `hermes-spreadsheet-suite`.
- Backlog này không thay thế [Missing Capabilities Backlog](./missing-capabilities-backlog-2026-04-23.md).
- Mục tiêu ở đây là tăng `reviewability`, `operability`, `UX ở host`, và `devex`, trong khi vẫn giữ nguyên các invariant hiện tại:
  - Hermes core không vào repo này
  - Step 1 / Step 2 contracts không bị churn tùy tiện
  - hosts vẫn là thin client
  - gateway vẫn là validator / approval / control plane

Nguồn ý tưởng chính:
- `asm` PR `#226`, `#220`, `#232`, `#231`, `#230`, `#218`, `#211`, `#208`, `#203`, `#200`, `#199`, `#196`, `#195`

## 1. Fixture Eval Runner Và Batch Regression Gate

**idea học từ `asm`**

- `#196`: batch-evaluate collections với provenance rõ ràng
- `#218`: eval -> fix -> re-eval loop với target threshold

**áp vào repo này như thế nào**

Tạo một `fixture-eval runner` cho các flow quan trọng của spreadsheet suite, thay vì chỉ dựa vào unit tests rời rạc.

Runner này nên đọc các fixture pack cố định và kiểm tra:
- request normalization có đúng contract không
- structured-body normalization có fail-closed đúng không
- reviewer-safe unavailable mode có giữ invariant không
- writeback preview / approval / completion có còn match plan không
- undo / redo có còn exact-safe trong các subset đang support không

**repo surfaces**

- `packages/contracts/src/`
- `services/gateway/src/hermes/`
- `services/gateway/src/routes/`
- `services/gateway/tests/`
- `apps/excel-addin/src/taskpane/`
- `apps/google-sheets-addon/src/`
- `scripts/`

**acceptance criteria**

- Có format fixture rõ ràng cho `excel`, `google_sheets`, và `gateway-only`
- Có lệnh kiểu `npm run eval:fixtures`
- Có output machine-readable để CI gate được
- Có aggregate summary theo capability family và host
- Có ít nhất các fixture chuẩn cho:
  - selection explain
  - formula help
  - screenshot unavailable-safe
  - screenshot real preview/import
  - writeback approval/completion
  - undo/redo exact-safe subset

**độ khó**

`High`

**thứ tự nên làm**

`1`

## 2. Host-State-Isolated Integration Harness

**idea học từ `asm`**

- `#200`: test phải tách khỏi host state thật

**áp vào repo này như thế nào**

Hiện Excel host, Google Sheets host, local storage, runtime config, session identity, và gateway polling đều là nguồn dễ làm test phụ thuộc môi trường.

Cần đưa các integration test quan trọng về trạng thái `deterministic`:
- không phụ thuộc Office runtime thật
- không phụ thuộc Apps Script runtime thật
- không phụ thuộc local storage state cũ
- không phụ thuộc auth / machine config / timing ngẫu nhiên

**repo surfaces**

- `apps/excel-addin/src/taskpane/taskpane.js`
- `apps/google-sheets-addon/html/Sidebar.js.html`
- `packages/shared-client/tests/`
- `services/gateway/tests/`

**acceptance criteria**

- Có fake harness cho Office.js / Apps Script / local storage / polling
- Các test host critical flow chạy được trên CI sạch, không cần runtime thật
- Không còn test nào pass/fail tùy vào session state cục bộ
- Có regression test cho:
  - attachment upload lifecycle
  - under-specified affirmations
  - undo / redo shortcut prompts
  - stale preview / stale approval handling

**độ khó**

`Medium`

**thứ tự nên làm**

`2`

## 3. Scannable Run History Và Two-Pane Explorer Cho Host UI

**idea học từ `asm`**

- `#195`: output lớn phải scannable
- `#231`: sidebar + detail layout
- `#232`: virtualize list khi số item lớn

**áp vào repo này như thế nào**

UI hiện tại chủ yếu là chat stream + message polling. Khi số lượng request, trace event, approval state, và execution history tăng lên, việc tìm lại một run cũ hoặc xem chính xác plan/result nào đã xảy ra sẽ ngày càng khó.

Nên thêm một `run explorer` trong host:
- cột trái: list các run / execution history entries
- cột phải: detail của run đang chọn
- mobile/narrow host: drawer mode

**repo surfaces**

- `apps/excel-addin/src/taskpane/taskpane.html`
- `apps/excel-addin/src/taskpane/taskpane.css`
- `apps/excel-addin/src/taskpane/taskpane.js`
- `apps/google-sheets-addon/html/Sidebar.html`
- `apps/google-sheets-addon/html/Sidebar.css.html`
- `apps/google-sheets-addon/html/Sidebar.js.html`
- `services/gateway/src/routes/executionControl.ts`
- `services/gateway/src/routes/trace.ts`

**acceptance criteria**

- Có list riêng cho recent runs / execution history thay vì chỉ chat scrollback
- Chọn một run sẽ thấy:
  - plan summary
  - preview / approval state
  - completion result
  - trace excerpt
  - undo / redo eligibility
- Danh sách dài được virtualize hoặc ít nhất window hóa
- State chọn run được giữ qua refresh nhẹ hoặc reopen taskpane

**độ khó**

`High`

**thứ tự nên làm**

`3`

## 4. Workflow Bundles / Demo Packs Thành First-Class Manifest

**idea học từ `asm`**

- `#211`: ship predefined bundles
- `#208`: bundle modify / export

**áp vào repo này như thế nào**

Repo này đã có:
- `docs/demo-runbook.md`
- `scripts/generate_demo_pack.py`
- nhiều capability family có thể demo theo scenario

Nhưng hiện chưa có một `manifest` chuẩn cho workflow pack hoặc demo pack. Cần biến các scenario lặp lại thành đối tượng first-class để:
- demo ổn định hơn
- QA chạy lại được
- docs không trôi khỏi test fixtures
- future sales/review packs dùng cùng một nguồn

**repo surfaces**

- `scripts/generate_demo_pack.py`
- `docs/demo-runbook.md`
- `docs/review/`
- `docs/setup/`
- thư mục mới kiểu `data/demo-packs/` hoặc `docs/demo-packs/`

**acceptance criteria**

- Có manifest format cho named scenarios
- Có ít nhất các pack chuẩn cho:
  - screenshot import preview
  - cleanup flow
  - range transfer
  - materialized analysis report
  - pivot/chart flow
- Có script export hoặc render pack thành checklist/demo assets
- Docs demo dùng cùng manifest thay vì mô tả tay riêng

**độ khó**

`Medium`

**thứ tự nên làm**

`4`

## 5. Generated Capability Registry: Slim Matrix + Detail Artifacts

**idea học từ `asm`**

- `#220`: split artifact theo use case runtime
- `#203`: id phải đủ specific để không collapse variants

**áp vào repo này như thế nào**

Capability knowledge hiện đang nằm rải ở:
- contracts
- planner/runtime rules
- writeback routes
- docs/capability-surface
- behavior khác nhau giữa Excel và Google Sheets

Nên tạo một generated capability registry để:
- host UI biết chính xác capability nào support ở host hiện tại
- docs không drift khỏi runtime
- future admin/review surfaces lazy-load detail được

**repo surfaces**

- `packages/contracts/src/`
- `services/gateway/src/hermes/runtimeRules.ts`
- `docs/capability-surface.md`
- `scripts/`

**acceptance criteria**

- Có artifact kiểu `capabilities.min.json` cho host/UI/docs nhẹ
- Có detail artifact hoặc generated markdown cho từng capability family
- ID capability phân biệt được host / operation / variant, không collapse Excel và Google Sheets vào một key mơ hồ
- `docs/capability-surface.md` được generate hoặc verify từ nguồn này

**độ khó**

`Medium-High`

**thứ tự nên làm**

`5`

## 6. Refactor Host UI Shell Mà Không Đụng Wire Contracts

**idea học từ `asm`**

- `#230`: rewrite UI layer nhưng giữ nguyên data contract

**áp vào repo này như thế nào**

`taskpane.js` và `Sidebar.js.html` đang khá lớn, đồng thời chứa cả:
- UI state
- runtime/session helpers
- request assembly
- approval/history logic
- rendering logic

Nếu tiếp tục mở rộng trực tiếp trong các file này, chi phí thay đổi sẽ tăng nhanh. Cần một backlog refactor dần host shell, nhưng phải giữ nguyên:
- request envelope
- response contract
- gateway routes
- review-safe invariants

**repo surfaces**

- `apps/excel-addin/src/taskpane/`
- `apps/google-sheets-addon/html/`
- `packages/shared-client/src/`

**acceptance criteria**

- View/state logic được tách module rõ hơn
- Shared client boundary được dùng nhiều hơn thay vì host tự cấy logic lặp
- Có parity checklist cho các flow hiện có trước và sau refactor
- Không có contract churn chỉ vì đổi UI shell

**độ khó**

`High`

**thứ tự nên làm**

`6`

## Không Nên Copy Nguyên Xi Lúc Này

- Bỏ Bun / gom toolchain về Node-only như `asm #226`
  - Repo này đã chủ yếu ở Node/TypeScript rồi; đây không phải choke point hiện tại.
- Website catalog rewrite như `asm #230`
  - Hermes Spreadsheet Suite hiện chưa có public catalog/product site kiểu đó; giá trị trực tiếp thấp hơn các mục trên.
- Public catalog filtering/highlighting như `asm #199`
  - Ý tưởng tốt, nhưng trước mắt nên áp nó vào run explorer nội bộ thay vì dựng website riêng.

## Khuyến nghị thứ tự tổng

1. Fixture eval runner và batch regression gate
2. Host-state-isolated integration harness
3. Scannable run history và two-pane explorer cho host UI
4. Workflow bundles / demo packs thành first-class manifest
5. Generated capability registry: slim matrix + detail artifacts
6. Refactor host UI shell mà không đụng wire contracts
