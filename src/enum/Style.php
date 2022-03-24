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
 * 样式代码
 *
 * Class Style
 * @package app\common\enum\html
 */
class Style
{
    /**
     * Length unit
     *
     * @const string
     */
    const UNIT_PT = 'pt'; // Mostly for shapes
    const UNIT_PX = 'px'; // Mostly for images

    /**
     * General positioning options.
     *
     * @const string
     */
    const POS_ABSOLUTE = 'absolute';
    const POS_RELATIVE = 'relative';

    /**
     * Horizontal/vertical value
     *
     * @const string
     */
    const POS_CENTER = 'center';
    const POS_LEFT = 'left';
    const POS_RIGHT = 'right';
    const POS_TOP = 'top';
    const POS_BOTTOM = 'bottom';
    const POS_INSIDE = 'inside';
    const POS_OUTSIDE = 'outside';

    /**
     * Position relative to
     *
     * @const string
     */
    const POS_RELTO_MARGIN = 'margin';
    const POS_RELTO_PAGE = 'page';
    const POS_RELTO_COLUMN = 'column'; // horizontal only
    const POS_RELTO_CHAR = 'char'; // horizontal only
    const POS_RELTO_TEXT = 'text'; // vertical only
    const POS_RELTO_LINE = 'line'; // vertical only
    const POS_RELTO_LMARGIN = 'left-margin-area'; // horizontal only
    const POS_RELTO_RMARGIN = 'right-margin-area'; // horizontal only
    const POS_RELTO_TMARGIN = 'top-margin-area'; // vertical only
    const POS_RELTO_BMARGIN = 'bottom-margin-area'; // vertical only
    const POS_RELTO_IMARGIN = 'inner-margin-area';
    const POS_RELTO_OMARGIN = 'outer-margin-area';

    /**
     * Wrap type
     *
     * @const string
     */
    const WRAP_INLINE = 'inline';
    const WRAP_SQUARE = 'square';
    const WRAP_TIGHT = 'tight';
    const WRAP_THROUGH = 'through';
    const WRAP_TOPBOTTOM = 'topAndBottom';
    const WRAP_BEHIND = 'behind';
    const WRAP_INFRONT = 'infront';
}