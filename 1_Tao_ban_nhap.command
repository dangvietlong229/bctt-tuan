#!/bin/zsh
cd "$(dirname "$0")" || exit 1
python3 weekly_report.py all
status=$?
echo
if [ $status -eq 0 ]; then
  echo "Đã tạo bản nháp. Hãy mở output_reports/latest_run.json để xem vị trí file review."
else
  echo "Có lỗi. Vui lòng xem thông báo phía trên."
fi
echo "Nhấn Enter để đóng cửa sổ."
read
exit $status
