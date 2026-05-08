# Office Local AutoSave

Autosave local thật cho Word/Excel Windows không cần Microsoft account/OneDrive.

Cơ chế: PowerShell watchdog chạy nền mỗi vài giây, attach vào instance Word/Excel đang mở qua COM, nếu file đã có path local và đang modified thì tạo backup rồi gọi Save.
Installer dùng `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` trỏ tới launcher VBScript chạy ẩn, không cần quyền admin và không bật cửa sổ PowerShell sau reboot.

## Cài đặt trên Windows

Cho người dùng lowtech: double-click `INSTALL_ONCE.cmd` một lần là xong. Sau reboot không cần chạy lại; watchdog tự chạy ẩn theo user Windows hiện tại.

Mở PowerShell bình thường, không cần admin:

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

Mặc định:

- Save mỗi 10 giây nếu file thay đổi.
- Backup snapshot tối đa mỗi 1 giờ trước khi save.
- Giữ backup 2 ngày.
- Tổng dung lượng backup tối đa 2GB.
- Luôn giữ trống tối thiểu 10GB trên ổ chứa backup; nếu không đủ thì bỏ qua backup nhưng vẫn save file gốc.
- Log: `%LOCALAPPDATA%\OfficeLocalAutoSave\autosave.log`
- Backup: `%LOCALAPPDATA%\OfficeLocalAutoSave\Backups`

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
3. Chờ 10-20 giây.
4. Mở log kiểm tra có dòng `saved`.
5. Kill Word/Excel từ Task Manager rồi mở lại file, nội dung phải còn.

Muốn đổi nhịp save/backup/quota:

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1 -IntervalSeconds 10 -BackupIntervalSeconds 3600 -BackupKeepDays 2 -BackupMaxMB 2048 -MinFreeSpaceMB 10240
powershell -ExecutionPolicy Bypass -File .\verify.ps1 -WaitSeconds 20
```

`IntervalSeconds` là nhịp kiểm tra/save file gốc. `BackupIntervalSeconds` là khoảng cách tối thiểu giữa 2 bản backup snapshot cho cùng một file; file gốc vẫn được save mỗi `IntervalSeconds` nếu có thay đổi. `BackupMaxMB` là quota tổng backup. `MinFreeSpaceMB` là mức dung lượng trống tối thiểu phải giữ lại.

## Lưu ý

- File mới chưa Save As thì không thể autosave vì chưa có đường dẫn.
- File readonly/protected/shared conflict sẽ bị bỏ qua.
- Đây là save vào file gốc, không phải AutoRecover.
