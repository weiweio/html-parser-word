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
 * tbl宽类型
 *
 * Class TblWidth
 * @package app\common\enum
 */
final class TblWidth
{
    //No Width
    const NIL = 'nil';

    //Automatically Determined Width
    const AUTO = 'auto';

    //Width in Fiftieths of a Percent
    const PERCENT = 'pct';

    //Width in Twentieths of a Point
    const TWIP = 'dxa';

    const VALUE = 8305;
}