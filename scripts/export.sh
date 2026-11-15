#!/bin/bash
# ส่งออก PDF ด้วย PowerPoint — อ้างไฟล์ตามชื่อ (ไม่ใช้ active presentation เพื่อไม่ไปปิดไฟล์อื่นที่ผู้ใช้เปิดอยู่)
# หมายเหตุ: ต้องรันนอก sandbox ของ Claude Code (ไม่งั้น PowerPoint เปิดไฟล์ไม่ได้ error -9074)
rm -f "$2"
N="$(basename "$1")"
osascript <<OSA
with timeout of 600 seconds
tell application "Microsoft PowerPoint"
  open POSIX file "$1"
  set p to presentation "$N"
  save p in POSIX file "$2" as save as PDF
  close p saving no
end tell
end timeout
OSA
