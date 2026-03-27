/**
 * HWP Studio Pre-Tool-Use Hook
 * HWPX 편집 규칙을 자동으로 검증합니다 (CLAUDE.md 규칙 적용).
 */
export default function preToolUse({ tool, input }) {
  // 1. fast-xml-parser 설치 차단 (CLAUDE.md 규칙 6)
  if (tool === 'Bash' && typeof input?.command === 'string') {
    if (input.command.includes('fast-xml-parser')) {
      return {
        decision: 'block',
        message: '[HWP 규칙 위반] fast-xml-parser 사용 금지.\n→ @xmldom/xmldom을 사용하세요.\n→ CLAUDE.md 규칙 6번 참조'
      };
    }
  }

  // 2. .tagName 사용 차단 → .localName 권장 (CLAUDE.md 규칙 7)
  if ((tool === 'Edit' || tool === 'Write') && typeof input?.new_string === 'string') {
    if (input.new_string.includes('.tagName')) {
      return {
        decision: 'block',
        message: '[HWP 규칙 위반] element.tagName 사용 금지.\n→ element.localName을 사용하세요.\n→ CLAUDE.md 규칙 7번 참조'
      };
    }
  }

  // 3. charPrIDRef 변경 차단 (CLAUDE.md 규칙 9)
  if (tool === 'Edit' && typeof input?.new_string === 'string') {
    if (input.new_string.includes('charPrIDRef') && input.file_path?.endsWith('.xml')) {
      return {
        decision: 'block',
        message: '[HWP 규칙 위반] charPrIDRef 변경 금지.\n→ 글자 속성 참조 ID를 변경하면 문서가 깨집니다.\n→ CLAUDE.md 규칙 9번 참조'
      };
    }
  }

  // 4. raw win32com 사용 차단 (CLAUDE.md 규칙 1)
  const writeContent = (tool === 'Write' && input?.content) || (tool === 'Edit' && input?.new_string) || '';
  if (writeContent && typeof writeContent === 'string') {
    if (writeContent.includes('import win32com') && !writeContent.includes('from pyhwpx')) {
      return {
        decision: 'block',
        message: '[HWP 규칙 위반] raw win32com 직접 사용 금지.\n→ from pyhwpx import Hwp를 사용하세요.\n→ CLAUDE.md 규칙 1번 참조'
      };
    }
  }
}
