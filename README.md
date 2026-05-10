# Office Local AutoSave

Autosave local thật cho Word/Excel Windows không cần Microsoft account/OneDrive.

Cơ chế: PowerShell watchdog chạy nền kiểm tra nhẹ mỗi giây, chỉ attach vào Word/Excel và autosave khi người dùng đã idle đủ lâu; nếu file đã có path local và đang modified thì tạo backup rồi gọi Save.
Installer dùng `HKLM\Software\Microsoft\Windows\CurrentVersion\Run` trỏ tới launcher VBScript chạy ẩn trong từng session user, cần quyền admin một lần và không bật cửa sổ PowerShell sau reboot.

## Cài đặt trên Windows

Cho người dùng lowtech: double-click `INSTALL_ONCE.cmd` một lần là xong. Windows sẽ hỏi quyền Administrator/UAC. Sau reboot không cần chạy lại; watchdog tự chạy ẩn cho mọi user Windows đăng nhập vào máy.

Mở PowerShell bình thường; installer sẽ tự xin quyền admin nếu cần:

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

Mặc định:

- Kiểm tra nhẹ mỗi 1 giây, chỉ Save nếu file thay đổi và người dùng idle tối thiểu 3 giây.
- Nếu Word/Excel đang là cửa sổ đang dùng, chỉ Save sau khi idle tối thiểu 10 giây để tránh đứng khi đang gõ.
- Backup snapshot tối đa mỗi 1 giờ trước khi save.
- Giữ backup 2 ngày.
- Tổng dung lượng backup tối đa 2GB.
- Luôn giữ trống tối thiểu 10GB trên ổ chứa backup; nếu không đủ thì bỏ qua backup nhưng vẫn save file gốc.
- Log: `%LOCALAPPDATA%\OfficeLocalAutoSave\autosave.log`
- Backup: `%LOCALAPPDATA%\OfficeLocalAutoSave\Backups`
- File cài đặt chung: `%ProgramData%\OfficeLocalAutoSave`

## Gỡ cài đặt

Double-click `UNINSTALL.cmd`, hoặc chạy:

```powershell
powershell -ExecutionPolicy Bypass -File .\uninstall.ps1
```

## Test nhanh

Chạy verify tự động:

Double-click `VERIFY.cmd`, hoặc chạy:

```powershell
powershell -ExecutionPolicy Bypass -File .\verify.ps1
```

Test thủ công:

1. Mở một file `.docx` hoặc `.xlsx` đã Save As local.
2. Sửa nội dung, không bấm Save.
3. Ngừng gõ/di chuột và chờ 15-25 giây.
4. Mở log kiểm tra có dòng `saved`.
5. Kill Word/Excel từ Task Manager rồi mở lại file, nội dung phải còn.

Muốn đổi nhịp save/backup/quota:

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1 -IntervalSeconds 1 -MinIdleSeconds 3 -ForegroundMinIdleSeconds 10 -BackupIntervalSeconds 3600 -BackupKeepDays 2 -BackupMaxMB 2048 -MinFreeSpaceMB 10240
powershell -ExecutionPolicy Bypass -File .\verify.ps1 -WaitSeconds 25
```

`IntervalSeconds` là nhịp kiểm tra an toàn nhẹ. `MinIdleSeconds` là thời gian tối thiểu không có input bàn phím/chuột trước khi watchdog được phép gọi Office Save. `ForegroundMinIdleSeconds` là thời gian idle tối thiểu khi Word/Excel đang là cửa sổ foreground, để tránh đứng khi đang gõ. File gốc được save ngay khi có thay đổi và đã idle đủ lâu, không ép Save đúng lúc người dùng đang thao tác. `BackupIntervalSeconds` là khoảng cách tối thiểu giữa 2 bản backup snapshot cho cùng một file. `BackupMaxMB` là quota tổng backup. `MinFreeSpaceMB` là mức dung lượng trống tối thiểu phải giữ lại.

## Lưu ý

- File mới chưa Save As thì không thể autosave vì chưa có đường dẫn.
- File readonly/protected/shared conflict sẽ bị bỏ qua.
- Đây là save vào file gốc, không phải AutoRecover.
