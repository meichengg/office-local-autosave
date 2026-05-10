# Changelog

## 2026-05-10

- Fix lag khi người dùng đang gõ bằng Mode A: watchdog chỉ kiểm tra nhẹ mỗi 1 giây và chỉ gọi Office COM Save khi người dùng đã idle đủ lâu.
- Đổi mặc định autosave an toàn: `IntervalSeconds=1`, `MinIdleSeconds=3`, `ForegroundMinIdleSeconds=10`.
- Đổi installer Windows sang machine-wide: cài vào `%ProgramData%\OfficeLocalAutoSave` và dùng `HKLM\Software\Microsoft\Windows\CurrentVersion\Run` để chạy ngầm cho mọi user đăng nhập.
- Thêm UAC elevation cho install/uninstall và dọn legacy `HKCU Run` entries từ bản current-user cũ.
- Cập nhật verify để kiểm tra `HKLM Run`, watchdog trong đúng session hiện tại, và chờ mặc định `25` giây.
- Cập nhật README và hướng dẫn người dùng tiếng Việt có dấu, nêu rõ chương trình không ép lưu đúng lúc đang gõ để tránh làm Word/Excel bị đứng khung.
