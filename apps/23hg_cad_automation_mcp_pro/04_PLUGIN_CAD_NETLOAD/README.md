# 04_PLUGIN_CAD_NETLOAD - KIẾN TRÚC KẾT NỐI 23HG

## 💎 NÂNG CẤP KIẾN TRÚC: NATIVE COM INTEROP (ZERO-PLUGIN)
Từ phiên bản 2.0.0, **23HG CAD AUTOMATION MCP PRO** chuyển đổi hoàn toàn sang kiến trúc **Native COM Interop** (AutoCAD ActiveX Automation):

1. **Không cần lệnh NETLOAD:** Người dùng chỉ cần mở AutoCAD, AI sẽ tự động bắt kết nối trực tiếp hai chiều qua Windows COM API.
2. **Loại bỏ hoàn toàn rủi ro DLL bên ngoài:**
   - Không lo Windows Smart App Control / Defender chặn file DLL chưa ký số.
   - Không lo xung đột phiên bản runtime (.NET Framework 4.8 vs .NET 8 / .NET 10).
3. **Tiện ích AutoLISP đi kèm:**
   File `23hg_bridge.lsp` cung cấp các lệnh hỗ trợ trong CAD (nạp qua lệnh `APPLOAD` nếu muốn dùng phím tắt).
