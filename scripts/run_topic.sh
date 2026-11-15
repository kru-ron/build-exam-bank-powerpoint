#!/bin/bash
# สร้างไฟล์ PowerPoint ของหัวข้อหนึ่ง แล้ววนวัดบรรทัดจริง + แก้วลีถูกตัด จนนิ่ง
# usage: run_topic.sh <data_module> <img_dir> <out.pptx>
#   ทำงานในโฟลเดอร์ที่มีไฟล์ data module (เช่น scratchpad ที่คัดลอก data/ มา)
#   ต้องมี Microsoft PowerPoint (macOS) สำหรับส่งออก PDF
SK="$(cd "$(dirname "$0")" && pwd)"
TPL="${TEMPLATE:-$SK/../assets/template.pptx}"
M=$1; I=$2; O=$3; P=${O%.pptx}.pdf
for i in 1 2 3 4 5 6 7 8 9 10; do
  python3 "$SK/build_bank.py" $M "$TPL" $I $O >/dev/null || exit 1
  "$SK/export.sh" "$PWD/$O" "$PWD/$P"
  python3 "$SK/measure.py" $P $O.keys.json /tmp/m_$M.json >/dev/null
  D=$(python3 -c "
import json,os; p='$O.lines.json'; a=json.load(open(p)) if os.path.exists(p) else {}; b=json.load(open('/tmp/m_$M.json'))
d=len([k for k,v in b.items() if a.get(k)!=v]); a.update(b); json.dump(a,open(p,'w'),ensure_ascii=False); print(d)")
  N=$(python3 "$SK/breaks.py" /tmp/m_$M.json.wrap.json $O.breaks.json | awk '{print $3}')
  echo "pass $i: line-diffs $D new-breaks $N"
  if [ "$D" = "0" ] && [ "$N" = "0" ]; then break; fi
done
python3 "$SK/build_bank.py" $M "$TPL" $I $O
"$SK/export.sh" "$PWD/$O" "$PWD/$P"
python3 "$SK/verify_deck.py" $O $P
