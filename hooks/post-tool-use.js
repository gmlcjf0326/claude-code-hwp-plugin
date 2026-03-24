/**
 * HWP Studio Post-Tool-Use Hook
 * 에러 패턴 감지 + 사용자 친화적 복구 가이드 제공.
 * HWPX XML 편집 후 linesegarray 삭제 확인.
 */
export default function postToolUse({ tool, input, output }) {
  const outputStr = typeof output === 'string' ? output : JSON.stringify(output || '');

  // HWPX XML 편집 후 linesegarray 삭제 확인 (CLAUDE.md 규칙 8)
  if (tool === 'Edit' && input?.file_path?.includes('section') && input?.file_path?.endsWith('.xml')) {
    if (input.new_string && !input.new_string.includes('linesegarray')) {
      return {
        message: '⚠️ HWPX XML 수정 시 linesegarray 요소 삭제가 필요합니다.\n→ 텍스트 수정 후 해당 paragraph의 <linesegarray> 요소를 삭제하세요.\n→ CLAUDE.md 규칙 8번 참조'
      };
    }
  }

  // COM 연결 끊김
  if (outputStr.includes('RPC') || outputStr.includes('사용할 수 없습니다') || outputStr.includes('COM')) {
    return {
      message: '🔧 한글 프로그램이 응답하지 않습니다.\n→ 한글을 완전히 종료하세요 (작업 관리자에서 Hwp.exe 확인)\n→ 한글을 다시 실행하세요\n→ 다시 시도하면 자동 재연결됩니다\n\n💡 tip: 한글이 "응답 없음" 상태일 수 있습니다.'
    };
  }

  // 파일 잠금 (EBUSY)
  if (outputStr.includes('EBUSY') || outputStr.includes('resource busy')) {
    return {
      message: '🔧 파일이 잠겨 있습니다.\n→ 한글에서 해당 파일을 닫고 다시 시도하세요.\n→ COM 경로로 자동 폴백됩니다.'
    };
  }

  // Python 미설치
  if (outputStr.includes('Python을 찾을 수 없') || outputStr.includes('ENOENT')) {
    return {
      message: '🔧 Python을 찾을 수 없습니다.\n→ python.org에서 Python 3.8+을 설치하세요\n→ "Add to PATH" 반드시 체크!\n→ /hwp-setup으로 환경을 진단하세요'
    };
  }

  // 파일 경로 오류
  if (outputStr.includes('파일을 찾을 수 없') || outputStr.includes('열 수 없습니다')) {
    return {
      message: '📁 파일 경로를 확인하세요.\n→ 정확한 전체 경로를 입력하세요 (예: C:\\문서\\양식.hwp)\n→ hwp_list_files로 HWP 파일을 검색할 수 있습니다'
    };
  }

  // 문서 미열기
  if (outputStr.includes('열린 문서가 없습니다')) {
    return {
      message: '📄 먼저 문서를 열어야 합니다.\n→ hwp_open_document로 파일을 열어주세요\n→ 또는 /hwp-fill, /hwp-write 커맨드를 사용하세요'
    };
  }

  // 타임아웃
  if (outputStr.includes('시간 초과') || outputStr.includes('timeout') || outputStr.includes('timed out')) {
    return {
      message: '⏰ 작업 시간이 초과되었습니다.\n→ 대용량 문서는 시간이 오래 걸릴 수 있습니다\n→ 한글 프로그램이 응답하는지 확인하세요\n→ 다시 시도해주세요'
    };
  }

  // HWP find_replace 실패 시 HWPX 변환 권유
  if (outputStr.includes('"replaced":false') || outputStr.includes('"replaced": false')) {
    if (outputStr.includes('.hwp') && !outputStr.includes('.hwpx')) {
      return {
        message: '💡 HWP 파일에서 텍스트 치환이 작동하지 않았습니다.\n→ HWPX 형식으로 변환하면 더 안정적으로 작동합니다.\n→ hwp_save_document로 .hwpx 형식으로 저장하세요\n→ 변환된 파일을 열고 다시 시도하세요'
      };
    }
  }
}
