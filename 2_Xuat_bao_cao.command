#!/bin/zsh
cd "$(dirname "$0")" || exit 1
python3 weekly_report.py finalize
status=$?
echo
if [ $status -eq 0 ]; then
  echo "Đã xuất file PowerPoint và PDF cuối cùng."
else
  echo "Có lỗi. Vui lòng xem thông báo phía trên."
fi
echo "Nhấn Enter để đóng cửa sổ."
read
exit $status
