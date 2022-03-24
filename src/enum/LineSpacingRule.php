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
 * Line Spacing Rule
 *
 * Class LineSpacingRule
 * @package app\common\enum
 */
final class LineSpacingRule
{
    /**
     * Automatically Determined Line Height
     */
    const AUTO = 'auto';

    /**
     * Exact Line Height
     */
    const EXACT = 'exact';

    /**
     * Minimum Line Height
     */
    const AT_LEAST = 'atLeast';
}
