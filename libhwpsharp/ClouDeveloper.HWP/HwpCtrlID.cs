using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClouDeveloper.HWP
{
    public static class HwpCtrlID
    {
        public static readonly uint CTRL_ID_SECTION_DEF = MAKE_CTRL_ID('s', 'e', 'c', 'd'); /* 구역 */
        public static readonly uint CTRL_ID_COLUMN_DEF = MAKE_CTRL_ID('c', 'o', 'l', 'd'); /* 단 */
        public static readonly uint CTRL_ID_HEADEDR = MAKE_CTRL_ID('h', 'e', 'a', 'd'); /* 머리말 */
        public static readonly uint CTRL_ID_FOOTER = MAKE_CTRL_ID('f', 'o', 'o', 't'); /* 꼬리말 */
        public static readonly uint CTRL_ID_FOOTNOTE = MAKE_CTRL_ID('f', 'n', ' ', ' '); /* 각주 */
        public static readonly uint CTRL_ID_ENDNOTE = MAKE_CTRL_ID('e', 'n', ' ', ' '); /* 미주 */
                                                                                        /* 자동 번호*/
        public static readonly uint CTRL_ID_AUTO_NUM = MAKE_CTRL_ID('a', 't', 'n', 'o');
        /* 새 번호 지정 */
        public static readonly uint CTRL_ID_NEW_NUM = MAKE_CTRL_ID('n', 'w', 'n', 'o');
        public static readonly uint CTRL_ID_PAGE_HIDE = MAKE_CTRL_ID('p', 'g', 'h', 'd'); /* 감추기 */
                                                                                          /* 페이지 번호 제어(97의 홀수쪽에서 시작) */
        public static readonly uint CTRL_ID_PAGE_NUM_CTRL = MAKE_CTRL_ID('p', 'g', 'c', 't');
        /* 쪽번호 위치 */
        public static readonly uint CTRL_ID_PAGE_NUM_POS = MAKE_CTRL_ID('p', 'g', 'n', 'p');
        /* 찾아보기 표식 */
        public static readonly uint CTRL_ID_INDEX_MARK = MAKE_CTRL_ID('i', 'd', 'x', 'm');
        public static readonly uint CTRL_ID_BOKM = MAKE_CTRL_ID('b', 'o', 'k', 'm'); /* 책갈피 */
                                                                                     /* 글자 겹침 */
        public static readonly uint CTRL_ID_TCPS = MAKE_CTRL_ID('t', 'c', 'p', 's');
        public static readonly uint CTRL_ID_DUTMAL = MAKE_CTRL_ID('t', 'd', 'u', 't'); /* 덧말 */
                                                                                       /* 숨은 설명 */
        public static readonly uint CTRL_ID_TCMT = MAKE_CTRL_ID('t', 'c', 'm', 't');

        public static readonly uint CTRL_ID_TABLE = MAKE_CTRL_ID('t', 'b', 'l', ' '); /* 표 */
        public static readonly uint CTRL_ID_LINE = MAKE_CTRL_ID('$', 'l', 'i', 'n'); /* 선 */
        public static readonly uint CTRL_ID_RECT = MAKE_CTRL_ID('$', 'r', 'e', 'c'); /* 사각형 */
        public static readonly uint CTRL_ID_ELL = MAKE_CTRL_ID('$', 'e', 'l', 'l'); /* 타원 */
        public static readonly uint CTRL_ID_ARC = MAKE_CTRL_ID('$', 'a', 'r', 'c'); /* 호 */
        public static readonly uint CTRL_ID_POLY = MAKE_CTRL_ID('$', 'p', 'o', 'l'); /* 다각형 */
        public static readonly uint CTRL_ID_CURV = MAKE_CTRL_ID('$', 'c', 'u', 'r'); /* 곡선 */
                                                                                     /* 한글97 수식 */
        public static readonly uint CTRL_ID_EQEDID = MAKE_CTRL_ID('e', 'q', 'e', 'd');
        public static readonly uint CTRL_ID_PIC = MAKE_CTRL_ID('$', 'p', 'i', 'c'); /* 그림 */
        public static readonly uint CTRL_ID_OLE = MAKE_CTRL_ID('$', 'o', 'l', 'e'); /* OLE */
                                                                                    /* 묶음 개체 */
        public static readonly uint CTRL_ID_CON = MAKE_CTRL_ID('$', 'c', 'o', 'n');

        public static readonly uint CTRL_ID_DRAWING_SHAPE_OBJECT = MAKE_CTRL_ID('g', 's', 'o', ' ');

        public static readonly uint CTRL_ID_FORM = MAKE_CTRL_ID('f', 'o', 'r', 'm');

        public static readonly uint FIELD_UNKNOWN = MAKE_CTRL_ID('%', 'u', 'n', 'k');
        /* 현재의 날짜/시간 필드 */
        public static readonly uint FIELD_DATE = MAKE_CTRL_ID('%', 'd', 't', 'e');
        /*파일 작성 날짜/시간 필드 */
        public static readonly uint FIELD_DOCDATE = MAKE_CTRL_ID('%', 'd', 'd', 't');
        /* 문서 경로 필드 */
        public static readonly uint FIELD_PATH = MAKE_CTRL_ID('%', 'p', 'a', 't');
        /*블럭 책갈피 */
        public static readonly uint FIELD_BOOKMARK = MAKE_CTRL_ID('%', 'b', 'm', 'k');
        /* 메일 머지 */
        public static readonly uint FIELD_MAILMERGE = MAKE_CTRL_ID('%', 'm', 'm', 'g');
        /* 상호 참조 */
        public static readonly uint FIELD_CROSSREF = MAKE_CTRL_ID('%', 'x', 'r', 'f');
        public static readonly uint FIELD_FORMULA = MAKE_CTRL_ID('%', 'f', 'm', 'u'); /* 계산식 */
        public static readonly uint FIELD_CLICKHERE = MAKE_CTRL_ID('%', 'c', 'l', 'k'); /* 누름틀 */
                                                                                        /* 문서 요약 정보 필드 */
        public static readonly uint FIELD_SUMMARY = MAKE_CTRL_ID('%', 's', 'm', 'r');
        /* 사용자 정보 필드 */
        public static readonly uint FIELD_USERINFO = MAKE_CTRL_ID('%', 'u', 's', 'r');
        /* 하이퍼링크 */
        public static readonly uint FIELD_HYPERLINK = MAKE_CTRL_ID('%', 'h', 'l', 'k');
        public static readonly uint FIELD_REVISION_SIGN = MAKE_CTRL_ID('%', 's', 'i', 'g');
        public static readonly uint FIELD_REVISION_DELETE = MAKE_CTRL_ID('%', '%', '*', 'd');
        public static readonly uint FIELD_REVISION_ATTACH = MAKE_CTRL_ID('%', '%', '*', 'a');
        public static readonly uint FIELD_REVISION_CLIPPING = MAKE_CTRL_ID('%', '%', '*', 'C');
        public static readonly uint FIELD_REVISION_SAWTOOTH = MAKE_CTRL_ID('%', '%', '*', 'S');
        public static readonly uint FIELD_REVISION_THINKING = MAKE_CTRL_ID('%', '%', '*', 'T');
        public static readonly uint FIELD_REVISION_PRAISE = MAKE_CTRL_ID('%', '%', '*', 'P');
        public static readonly uint FIELD_REVISION_LINE = MAKE_CTRL_ID('%', '%', '*', 'L');
        public static readonly uint FIELD_REVISION_SIMPLECHANGE = MAKE_CTRL_ID('%', '%', '*', 'c');
        public static readonly uint FIELD_REVISION_HYPERLINK = MAKE_CTRL_ID('%', '%', '*', 'h');
        public static readonly uint FIELD_REVISION_LINEATTACH = MAKE_CTRL_ID('%', '%', '*', 'A');
        public static readonly uint FIELD_REVISION_LINELINK = MAKE_CTRL_ID('%', '%', '*', 'i');
        public static readonly uint FIELD_REVISION_LINETRANSFER = MAKE_CTRL_ID('%', '%', '*', 't');
        public static readonly uint FIELD_REVISION_RIGHTMOVE = MAKE_CTRL_ID('%', '%', '*', 'r');
        public static readonly uint FIELD_REVISION_LEFTMOVE = MAKE_CTRL_ID('%', '%', '*', 'l');
        public static readonly uint FIELD_REVISION_TRANSFER = MAKE_CTRL_ID('%', '%', '*', 'n');
        public static readonly uint FIELD_REVISION_SIMPLEINSERT = MAKE_CTRL_ID('%', '%', '*', 'e');
        public static readonly uint FIELD_REVISION_SPLIT = MAKE_CTRL_ID('%', 's', 'p', 'l');
        public static readonly uint FIELD_REVISION_CHANGE = MAKE_CTRL_ID('%', '%', 'm', 'r');
        public static readonly uint FIELD_MEMO = MAKE_CTRL_ID('%', '%', 'm', 'e');
        public static readonly uint FIELD_PRIVATE_INFO_SECURITY = MAKE_CTRL_ID('%', 'c', 'p', 'r');

        public static uint MAKE_CTRL_ID(char a, char b, char c, char d)
        {
            return (uint)((((byte)(a)) << 24) |
                (((byte)(b)) << 16) |
                (((byte)(c)) << 8) |
                (((byte)(d)) << 0));
        }
    }
}
