#!/usr/bin/env bash
# 새 항목 만들기.
#   ./new.sh exp  supabase-realtime   -> experiments/2026-09-13-supabase-realtime/README.md
#   ./new.sh proj my-app              -> projects/my-app.md
#   ./new.sh note vercel-cron         -> notes/vercel-cron.md
set -euo pipefail

cd "$(dirname "$0")"

kind=${1:-}
slug=${2:-}

if [ -z "$kind" ] || [ -z "$slug" ]; then
  echo "사용법: ./new.sh {exp|proj|note} <이름>" >&2
  exit 1
fi

today=$(date +%F)

case "$kind" in
  exp)
    dest="experiments/${today}-${slug}"
    [ -e "$dest" ] && { echo "이미 있음: $dest" >&2; exit 1; }
    mkdir -p "$dest"
    sed -e "s/<실험 이름>/${slug}/" -e "s/YYYY-MM-DD/${today}/g" \
      templates/experiment.md > "$dest/README.md"
    out="$dest/README.md"
    ;;
  proj)
    out="projects/${slug}.md"
    [ -e "$out" ] && { echo "이미 있음: $out" >&2; exit 1; }
    sed -e "s/<프로젝트 이름>/${slug}/" -e "s/<repo>/${slug}/" -e "s/YYYY-MM-DD/${today}/g" \
      templates/project.md > "$out"
    ;;
  note)
    out="notes/${slug}.md"
    [ -e "$out" ] && { echo "이미 있음: $out" >&2; exit 1; }
    sed -e "s/<주제>/${slug}/" -e "s/YYYY-MM-DD/${today}/g" \
      templates/note.md > "$out"
    ;;
  *)
    echo "알 수 없는 종류: $kind (exp|proj|note 중 하나)" >&2
    exit 1
    ;;
esac

echo "만들었습니다: $out"
echo "→ 내용을 채우고 README.md 표에 한 줄 추가하세요."
