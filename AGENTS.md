# AGENTS.md — 23HG MULTI-AGENT SYSTEM (AEC & CIVIL ENGINEERING)

Hướng dẫn vận hành chuẩn cho AI Coding Agents (Antigravity, Gemini, Claude, Cursor) tại **23HG Multi-Agent System**.

> **Nguồn sự thật tối cao (Single Source of Truth):** Mọi quy ước về lệnh chạy test, nguyên tắc số liệu, trạng thái pháp lý và vệ sinh kho mã nguồn được quy định tại [`CLAUDE.md`](./CLAUDE.md) và [`workflows/LEGAL_STATUS.md`](./workflows/LEGAL_STATUS.md). File `AGENTS.md` này kết nối các quy ước đó với bộ kỹ năng kỹ nghệ phần mềm (`agent-skills`), tuyệt đối không ghi đè hay làm yếu đi [`CLAUDE.md`](./CLAUDE.md).

---

## 1. Kiến Trúc Kết Hợp 2 Tầng (2-Layer Architecture)
- **Tầng 1 — Nghiệp vụ AEC tất định (Python Core & MCP):**
  - Không dùng LLM trong lõi tính toán: đo bóc tiên lượng, cắt thép 1D CSP, dự toán $G_{XD}$, thanh toán 03a, tiến độ CPM, KCS, As-Built.
  - Tình trạng hiệu lực pháp lý của các Thông tư / Nghị định: bắt buộc tra cứu tại [`workflows/LEGAL_STATUS.md`](./workflows/LEGAL_STATUS.md) (không tự ghi "đã đối chiếu bản gốc" khi chưa có nguồn xác minh).
- **Tầng 2 — Kỷ luật Kỹ nghệ Phần mềm (`agent-skills` — 12 kỹ năng cốt lõi):**
  - `interview-me`: Phỏng vấn làm rõ yêu cầu (từng câu một kèm dự đoán) khi thông số dự án còn thiếu.
  - `spec-driven-development` & `planning-and-task-breakdown`: Lập đặc tả và chia lát cắt dọc nguyên tử trước khi code.
  - `constraint-driven-development`: Thực thi khế ước chất lượng cơ học tại [`CONSTRAINTS.md`](./CONSTRAINTS.md).
  - `test-driven-development`: Viết test chứng minh lỗi/thiếu sót (RED) trước khi viết code (GREEN).
  - `debugging-and-error-recovery`: Cô lập nguyên nhân gốc rễ theo 5 bước, cấm đoán mò hoặc nuốt lỗi im lặng.
  - `doubt-driven-development` & `code-review-and-quality`: Phản biện chéo và thẩm tra 5 trục trước khi bàn giao.
  - `code-simplification` & `context-engineering`: Tinh gọn mã nguồn, khử bẫy Modal Dialog, tiết kiệm token.
  - `security-and-hardening` & `git-workflow-and-versioning`: Bảo vệ đường dẫn cá nhân và vệ sinh Git.

---

## 2. Quy Tắc Bắt Buộc Trước Khi Kết Thúc Mọi Phiên Làm Việc
1. Chạy đầy đủ các lệnh kiểm tra cơ học trong [`CONSTRAINTS.md`](./CONSTRAINTS.md).
2. Không để lại script nháp hoặc đường dẫn máy cá nhân trong `examples/` hoặc thư mục gốc.
