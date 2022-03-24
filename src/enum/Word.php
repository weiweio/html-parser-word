<?php
/*
 * This file is part of the zmh/html-parser-word.
 *
 * (c) zhangmenghua <993187039@qq.com>
 *
 * This source file is subject to the MIT license that is bundled.
 */

namespace HtmlToWord\enum;


/**
 * word样式
 * Class Word
 * @package HtmlToWord\enum
 */
class Word
{
    const WP_HEADING = '<w:p><w:pPr><w:pStyle w:val="#"/></w:pPr>';

    const WP_ALIGN = '<w:p><w:pPr><w:jc w:val="#"/></w:pPr>';

    const WP = '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/><w:textAlignment w:val="center"/><w:ind w:left="273" w:hanging="273" w:hangingChars="130"/></w:pPr>';

    const WP_OPTION = '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/><w:ind w:firstLine="273" w:firstLineChars="130"/><w:jc w:val="left"/><w:textAlignment w:val="center"/></w:pPr>';

    const WP_END="</w:p>";

    const TBL_PR = '<w:tblPr><w:tblW w:w="#width#" w:type="#unit#"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders><w:tblLayout w:type="autofit"/></w:tblPr>';

    const RELATIONSHIP ="<Relationship Id=\"rId#117#\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/#image.wmf#\"/>";

    const DOCUMENT_PIC_MARK = "<w:r><w:pict><v:shape id=\"_x0000#_i1051#\" type=\"_x0000_t75\" style=\"##\"><v:imagedata r:id=\"rId#27#\" o:title=\"2\"/></v:shape></w:pict></w:r>";

    const PKG_PART ="<pkg:part pkg:name=\"/word/media/#image1.wmf#\" pkg:contentType=\"image/#x-wmf#\"><pkg:binaryData>#ACXCADSQWWDD#</pkg:binaryData></pkg:part>";

    const BR = '<w:r><w:br/></w:r>';

    const WR_START = '<w:r>';

    const WR_END = '</w:r>';

    const TBL_START = '<w:tbl>';

    const WTR_START = '<w:tr>';

    const WTC_START = '<w:tc>';

    const WTC_END = '</w:tc>';

    const WTR_END = '</w:tr>';

    const TBL_END = '</w:tbl>';

    const RPR_END = '</w:rPr>';

}