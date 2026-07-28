import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

def convert_doc_to_docx_with_word():
    """
    MS Word COM 인터페이스를 사용하여 .doc를 .docx로 변환한다.
    """
    try:
        import win32com.client
        print("✔ win32com 모듈 로드 성공")
    except ImportError:
        print("❌ win32com을 로드할 수 없습니다. pywin32 설치 필요")
        return False

    doc_file = r"d:\Dev\Project\sample\doc\정보통신설비_2026성능점검결과서_엠코지니어스타.doc"
    docx_file = r"d:\Dev\Project\sample\doc\정보통신설비_2026성능점검결과서_엠코지니어스타.docx"

    if not os.path.exists(doc_file):
        print(f"❌ 파일이 없습니다: {doc_file}")
        return False

    try:
        print("🔄 MS Word 실행 중...")
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False

        # DOC 파일 열기
        print(f"📖 파일 열기: {doc_file}")
        doc = word.Documents.Open(os.path.abspath(doc_file))

        # DOCX로 저장
        print(f"💾 DOCX로 저장: {docx_file}")
        doc.SaveAs(os.path.abspath(docx_file), FileFormat=16)  # 16 = wdFormatDocx

        doc.Close()
        word.Quit()

        print(f"✅ 변환 완료: {docx_file}")
        return True

    except Exception as e:
        print(f"❌ 변환 실패: {e}")
        return False

if __name__ == '__main__':
    success = convert_doc_to_docx_with_word()
    sys.exit(0 if success else 1)
