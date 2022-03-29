<?php
/*
 * This file is part of the zmh/html-parser-word.
 *
 * (c) zhangmenghua <993187039@qq.com>
 *
 * This source file is subject to the MIT license that is bundled.
 */

namespace HtmlToWord;

use DOMXPath;
use DOMDocument;
use XSLTProcessor;
use HtmlToWord\enum\Jc;
use HtmlToWord\enum\Word;
use HtmlToWord\enum\Style;
use HtmlToWord\enum\TblWidth;
use HtmlToWord\enum\LineSpacingRule;

/**
 * html文本解析方法
 *
 * Class Parse
 * @package src
 */
class Parse
{
    /**
     * @var string 主体内容
     */
    private $body;

    /**
     * @var integer 当前math标签索引值
     */
    private $mathIndex;

    /**
     * @var array mathml列表
     */
    private $mathList;

    /**
     * @var bool $isPStart 段落是否开启
     */
    private $isPStart = false;

    /**
     * 是否存在单独的段落
     * @var bool
     */
    private $isExistAloneP = false;

    /**
     * @var bool $continueEach 是否持续遍历
     */
    private $continueEach = true;


    /**
     * @var boolean 段落第一个标签不用换行
     */
    private $enableBreak = true;

    /**
     * @var array $relationshipStr 图片地址关系
     */
    private $relationshipStr = [];

    /**
     * @var array $pkgpartStr 图片地址关系
     */
    private $pkgpartStr = [];

    /* @var $options array 配置项 */
    private static $options;

    /** @var \DOMXPath $xpath */
    protected static $xpath;




    /**
     * html2word执行
     *
     * @param $html
     * @param bool $fullHTML
     * @param bool $preserveWhiteSpace
     * @param null $options
     */
    public function doHtml2Word($html, $fullHTML = false, $preserveWhiteSpace = true, $options = null)
    {
        //附加项
        self::$options = $options;
        //检查标签及过滤
        $html = self::check($html);
        if (false === $fullHTML) {
            $html = '<body>' . $html . '</body>';
        }
        //正则匹配math标签
        preg_match_all('/<math.*\/math>/iUs', $html, $result);
        if ($result[0]) {
            $this->mathList = $result[0];
        }
        //初始化值
        $this->body = '';
        $this->continueEach = true;
        $this->isPStart = false;
        $this->mathIndex = 0;
        $this->enableBreak = false;
        $this->isExistAloneP = false;
        // Load DOM
        $orignalLibEntityLoader = libxml_disable_entity_loader(true);
        $dom = new DOMDocument();
        $dom->preserveWhiteSpace = $preserveWhiteSpace;
        $dom->loadXML($html);
        //self::$xpath = new DOMXPath($dom);
        $node = $dom->getElementsByTagName('body');
        $this->parseWord($node->item(0));
        //追加结束段落标签
        if ($this->isPStart) {
            $this->body .= Word::WP_END;
        }
        libxml_disable_entity_loader($orignalLibEntityLoader);
    }




    /**
     * 检查html标签
     *
     * @param $html
     * @return string|string[]
     */
    protected static function check($html)
    {
        //替换元素
        $html = str_replace('&nbsp;', ' ', $html);
        $html = str_replace(array("\n", "\r"), '', $html);
        $html = str_replace(array('&lt;', '&gt;', '&amp;'), array('_lt_', '_gt_', '_amp_'), $html);
        $html = html_entity_decode($html, ENT_QUOTES, 'UTF-8');
        $html = str_replace('&', '&amp;', $html);
        $html = str_replace(array('_lt_', '_gt_', '_amp_'), array('&lt;', '&gt;', '&amp;'), $html);
        $html = str_replace('<br>', '<br />', $html);
        //math标签中将指定的特殊字符转换为HTML实体
        $html = str_replace('><<', '>&lt;<', $html);
        //检查img标签是否闭合
        preg_match_all('/<img([^\/].*?)src=\"(.*?)\"(.*?)>/is', $html, $match);
        foreach ($match[0] as $img) {

            //追加标签结束符
            if (strpos($img, '/>') === false) {
                //替换成
                $image = explode('>', $img)[0] . ' />';
                $html = substr_replace($html, $image, strpos($html, $img), strlen($img));
            }
        }
        return $html;
    }



    /**
     * html解析
     *
     * @param $node
     * @param array $styles
     * @param array $data
     */
    public function parseWord($node, $styles = [], $data = [])
    {
        //样式数据
        $styleTypes = array('font', 'paragraph', 'list', 'table');
        foreach ($styleTypes as $styleType) {
            if (!isset($styles[$styleType])) {
                $styles[$styleType] = array();
            }
        }

        //匹配节点
        $nodes = array(
            // $method     $node     $styles     $data   $argument1      $argument2
            'p'         => array('Paragraph',   $node,     $styles,    null,    null,          null),
            'h1'        => array('Heading',     null,      $styles,    null,   'Heading1',     null),
            'h2'        => array('Heading',     null,      $styles,    null,   'Heading2',     null),
            'h3'        => array('Heading',     null,      $styles,    null,   'Heading3',     null),
            'h4'        => array('Heading',     null,      $styles,    null,   'Heading4',     null),
            'h5'        => array('Heading',     null,      $styles,    null,   'Heading5',     null),
            'h6'        => array('Heading',     null,      $styles,    null,   'Heading6',     null),
            '#text'     => array('Text',        $node,     $styles,    null,    null,          null),
            'strong'    => array('Property',    null,      $styles,    null,   'bold',         true),
            'b'         => array('Property',    null,      $styles,    null,   'bold',         true),
            'em'        => array('Property',    null,      $styles,    null,   'italic',       true),
            'i'         => array('Property',    null,      $styles,    null,   'italic',       true),
            'u'         => array('Property',    null,      $styles,    null,   'underline',    'single'),
            'sup'       => array('Property',    null,      $styles,    null,   'superScript',  true),
            'sub'       => array('Property',    null,      $styles,    null,   'subScript',    true),
            'span'      => array('Span',        $node,     $styles,    null,    null,          null),
            'font'      => array('Span',        $node,     $styles,    null,    null,          null),
            'table'     => array('Table',       $node,     $styles,    null,    null,          null),
            'ul'        => array('List',        $node,     $styles,    null,    null,          null),
            'ol'        => array('List',        $node,     $styles,    null,    null,          null),
            'img'       => array('Image',       $node,     $styles,    null,    null,          null),
            'br'        => array('LineBreak',   null,      $styles,    null,    null,          null),
            'a'         => array('Link',        $node,     $styles,    null,    null,          null),
            'math'      => array('Math',        $node,     $styles,    null,    null,          null),
            'bdo'       => array('Bdo',         $node,     $styles,    null,    null,          null),
        );
        $keys = array('node', 'styles', 'data', 'argument1', 'argument2');
        if (isset($nodes[$node->nodeName])) {
            //解析
            $arguments = array();
            $args = array();
            list($method, $args[0], $args[1], $args[2], $args[3], $args[4]) = $nodes[$node->nodeName];
            for ($i = 0; $i <= 4; $i++) {
                if ($args[$i] !== null) {
                    $arguments[$keys[$i]] = &$args[$i];
                }
            }
            //执行解析方法
            $method = "parse{$method}";
            call_user_func_array(array('HtmlToWord\Parse', $method), $arguments);
            // 检索变量数据
            foreach ($keys as $key) {
                if (array_key_exists($key, $arguments)) {
                    $$key = $arguments[$key];
                }
            }
        }
        //遍历子集
        $cNodes = $node->childNodes;
        if ($this->continueEach && !empty($cNodes)) {
            foreach ($cNodes as $cNode) {
                $this->parseWord($cNode, $styles, $data);
            }
        }
        //继续迭代
        $this->continueEach = true;
    }




    /**
     * 解析段落节点
     *
     * @param \DOMNode $node
     * @param array &$styles
     */
    protected function parseParagraph($node, &$styles)
    {
        $styles['paragraph'] = self::recursiveParseStylesInHierarchy($node, $styles['paragraph']);
        //是否开启换行
        if ($this->enableBreak) {
            $this->body .= Word::BR;
        }
        !$this->enableBreak && $this->enableBreak = true;
    }




    /**
     * 解析文本
     *
     * @param \DOMNode $node
     * @param array &$styles
     */
    protected function parseText($node, &$styles)
    {
        $value = trim($node->nodeValue);
        if (!empty($value) || $value == 0) {
            $styles['font'] = self::recursiveParseStylesInHierarchy($node, $styles['font']);
            //设置段落样式
            if (isset($styles['font']['alignment']) && is_array($styles['paragraph'])) {
                $styles['paragraph']['alignment'] = $styles['font']['alignment'];
            }
            //wr开始
            $wr = Word::WR_START . '<w:rPr>';
            //添加文本
            if ($styles['font']) {
                if (isset($styles['font']['superScript'])) {
                    $wr .= '<w:vertAlign w:val="superscript"/>';
                }
                if (isset($styles['font']['subScript'])) {
                    $wr .= '<w:vertAlign w:val="subscript"/>';
                }
                if (isset($styles['font']['bold'])) {
                    $wr .= '<w:b w:val="true"/>';
                }
                if (isset($styles['font']['italic'])) {
                    $wr .= '<w:em w:val="' . $styles['font']['italic'] . '"/>';
                }
                if (isset($styles['font']['underline'])) {
                    $wr .= '<w:u w:val="' . $styles['font']['underline'] . '"/>';
                }
            }
            //wrpr结束
            $wr .= Word::RPR_END;
            //替换特殊字符
            $value = str_replace(array('<', '>'), array('&lt;', '&gt;'), $value);
            $wr .= "<w:t>{$value}</w:t>";
            //wr结束
            $wr .= Word::WR_END;
            //创建段落
            $this->createParagraph($styles);
            //连接段落文本
            $this->body .= $wr;
        }
    }


    /**
     * 解析行内样式
     *
     * @param $node
     * @param array $styles
     * @return array|mixed
     */
    protected static function parseInlineStyle($node, $styles = array())
    {
        if (XML_ELEMENT_NODE == $node->nodeType) {
            $attributes = $node->attributes; // get all the attributes(eg: id, class)

            foreach ($attributes as $attribute) {
                switch ($attribute->name) {
                    case 'style':
                        $styles = self::parseStyle($attribute, $styles);
                        break;
                    case 'align':
                        $styles['alignment'] = self::mapAlign($attribute->value);
                        break;
                    case 'lang':
                        $styles['lang'] = $attribute->value;
                        break;
                    case 'class':
                        $styles = self::parseClass($attribute, $styles);
                        break;
                }
            }
        }

        return $styles;
    }



    /**
     * 解析math标签
     *
     * @param \DOMNode $node
     * @param $styles
     */
    protected function parseMath($node, $styles)
    {
        //匹配的math
        $math = isset($this->mathList[$this->mathIndex])
            ? $this->mathList[$this->mathIndex] : '';
        if ($math) {
            //创建段落
            $this->createParagraph($styles);
            //mathml标签转成omml
            $omml = self::math2omml($math);
            $this->body .= $omml;
        }
        $this->mathIndex++;
        $this->continueEach = false;
    }


    /**
     * mathml转换成omml
     * @param $mathml
     * @return false|string|string[]
     */
    private static function math2omml($mathml)
    {
        /* @var XSLTProcessor $processor */
        static $processor;
        if (empty($processor)) {
            libxml_disable_entity_loader(false);
            $xsl        = new DOMDocument();
            $str = __DIR__ . '/../mml2omml.xsl';
            $xsl->load($str);
            $processor = new XSLTProcessor();
            $processor->importStyleSheet($xsl);
        }
        //加载mathml标签
        $domDocument = new DOMDocument();
        //替换命名空间
        $mathml = str_replace(":mml", '', $mathml);
        $domDocument->loadXML($mathml);
        $omml = $processor->transformToXml($domDocument);
        $omml = str_replace('<?xml version="1.0" encoding="UTF-8"?>', '', $omml);
        $omml = str_replace('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"', '', $omml);
        $omml = str_replace(":mml", '', $omml);
        return $omml;
    }


    /**
     * 解析bdo
     *
     * @param \DOMNode $node
     * @param $styles
     */
    protected function parseBdo($node, &$styles)
    {
        $style = self::parseInlineStyle($node, $styles['font']);
        $styles['font'] = $style;
        //if (self::$xpath->query('./br', $node)->length > 0) {
        //}
    }



    /**
     * 解析【h1|2|3|4|5|6】
     *
     * @param $element
     * @param $styles
     * @param $argument1
     * @return mixed
     */
    protected static function parseHeading(&$styles, $argument1)
    {
        $styles['paragraph'] = $argument1;
    }



    /**
     * Parse property node
     *
     * @param array &$styles
     * @param string $argument1 Style name
     * @param string $argument2 Style value
     */
    protected static function parseProperty(&$styles, $argument1, $argument2)
    {
        $styles['font'][$argument1] = $argument2;
    }


    /**
     * Parse span node
     *
     * @param \DOMNode $node
     * @param array &$styles
     */
    protected static function parseSpan($node, &$styles)
    {
        $style = self::parseInlineStyle($node, $styles['font']);
        $styles['font'] = $style;
        //if (isset($styles['font']['underline']) && $node->childNodes->length <= 0) {
        //}
    }



    /**
     * 解析表格元素
     *
     * @param \DOMNode $node
     * @param &$styles
     * @return bool
     */
    protected function parseTable($node, &$styles)
    {
        $elementStyles = self::parseInlineStyle($node, $styles['table']);
        //遍历table节点
        $nodes = $node->childNodes;
        if ($nodes->length < 1) {
            $this->continueEach = false;
            return false;
        }
        //检查是否包含tbody
        foreach ($nodes as $nod) {
            if ($nod->nodeName === 'tbody') {
                $nodes = $nod->childNodes;
                if ($nodes->length < 1) {
                    $this->continueEach = false;
                    return false;
                }
                break;
            }
        }
        //todo 遇上table自动提升一个段落
        if ($this->body) {
            //表示段落未结束
            if ($this->isPStart) {
                $this->body .= Word::WP_END;
                $this->isExistAloneP && $this->isExistAloneP = false;
                $this->isPStart = false;
            }
        }
        //表格开始
        $tbl = Word::TBL_START;
        //获取表格宽及单位
        if (isset($elementStyles['width'])) {
            $width = $elementStyles['width'];
            $unit = $elementStyles['unit'];
        } else {
            $width = TblWidth::VALUE;
            $unit = TblWidth::TWIP;
        }
        //表格属性
        $tblPr = str_replace('#width#', $width, Word::TBL_PR);
        $tblPr = str_replace('#unit#', $unit, $tblPr);
        $tbl .= $tblPr;
        $this->body .= $tbl;
        //计算td的数量
        $tdNum = 0;
        foreach ($nodes as $item) {
            if ($item->nodeName == 'tr') {
                foreach ($item->childNodes as $childNodes) {
                    if ($childNodes->nodeName == 'td') {
                        $tdNum++;
                    }
                }
                break;
            }
        }
        $gridCols = [];
        $gridColWidth = $width / $tdNum;
        for ($i=0; $i < $tdNum; $i++) {
            $gridCols[] = '<w:gridCol w:w="'.$gridColWidth.'"/>';
        }
        $this->body .= '<w:tblGrid>'.implode('', $gridCols).'</w:tblGrid>';
        //td样式
        $tcPr = '<w:tcPr/>';
        //定义段落样式
        $styles['paragraph']['alignment'] = Jc::CENTER;
        //遍历tr
        foreach ($nodes as $trNode) {
            if ($trNode->nodeName != 'tr') {
                continue;
            }
            //获取td
            $tdNodes = $trNode->childNodes;
            if ($tdNodes->length < 1) {
                continue;
            }
            //tr开始
            $this->body .= Word::WTR_START;
            //遍历td
            foreach ($tdNodes as $tdNode) {
                //判断是否是td
                if ($tdNode->nodeName != 'td') {
                    continue;
                }
                //td开始
                $this->body .= Word::WTC_START;
                //td样式
                //$styles = self::recursiveParseStylesInHierarchy($tdNode, []);
                $this->body .= $tcPr;
                //解析td子类
                if ($tdNode->childNodes->length > 0) {
                    //td段落创建
                    $this->createParagraph($styles);
                    //遍历td节点
                    foreach ($tdNode->childNodes as $childNode) {
                        $this->parseWord($childNode, $styles);
                    }
                    //关闭段落
                    if ($this->isPStart) {
                        $this->body .= Word::WP_END;
                        $this->isPStart = false;
                        $this->isExistAloneP = false;
                    }
                }
                //td结束
                $this->body .= Word::WTC_END;
            }
            //tr结束
            $this->body .= Word::WTR_END;
        }
        $this->body .= Word::TBL_END;
        //不在继续迭代
        $this->continueEach = false;
    }



    /**
     * 在父级节点上解析样式
     *
     * @param \DOMNode $node
     * @param array &$style
     * @return array
     */
    protected static function recursiveParseStylesInHierarchy(\DOMNode $node, array $style)
    {
        $parentStyle = self::parseInlineStyle($node, array());
        $style = array_merge($parentStyle, $style);
        if ($node->parentNode != null && XML_ELEMENT_NODE == $node->parentNode->nodeType) {
            $style = self::recursiveParseStylesInHierarchy($node->parentNode, $style);
        }

        return $style;
    }



    /**
     * 创建段落
     *
     * @param $styles
     */
    private function createParagraph($styles)
    {
        //表示已开启段落
        if ($this->isPStart) {
            //如果开启了单独段落但是未设置段落样式就开启正常段落
            if ($this->isExistAloneP && empty($styles['paragraph'])) {
                $this->body .= Word::WP_END . (isset(self::$options['is_option']) ? Word::WP_OPTION : Word::WP);
                $this->isExistAloneP = false;
            }elseif($styles['paragraph'] && !$this->isExistAloneP) {
                //如果存在段落样式但是未开启单独段落则开启单独段落
                if (isset($styles['paragraph']['alignment'])) {
                    //段落开始
                    $wp = str_replace('#', $styles['paragraph']['alignment'], Word::WP_ALIGN);
                    $this->body .= Word::WP_END . $wp;
                    $this->isExistAloneP = true;
                }elseif (is_string($styles['paragraph'])) {
                    //表示h1|2|3|4|5|6主标题
                    $wp = str_replace('#', $styles['paragraph'], Word::WP_HEADING);
                    $this->body .= Word::WP_END . $wp;
                    $this->isExistAloneP = true;
                }
            }
        }else {
            if (isset($styles['paragraph']['alignment'])) {
                $wp = str_replace('#', $styles['paragraph']['alignment'], Word::WP_ALIGN);
                $this->isExistAloneP = true;
            }elseif (is_string($styles['paragraph'])) {
                //表示h1|2|3|4|5|6主标题
                $wp = str_replace('#', $styles['paragraph'], Word::WP_HEADING);
                $this->isExistAloneP = true;
            }else {
                $wp = isset(self::$options['is_option']) ? Word::WP_OPTION : Word::WP;
            }
            $this->body .= $wp;
            $this->isPStart = true;
        }
    }



    /**
     * 解析ul、ol
     *
     * @param \DOMNode $node
     * @param array &$styles
     */
    protected function parseList($node, &$styles)
    {
        if ($node->childNodes->count() > 0) {
            //todo 遇上ul、ol自动提升一个段落
            if ($this->body) {
                //表示段落未结束
                if ($this->isPStart) {
                    $this->body .= Word::WP_END;
                    $this->isExistAloneP && $this->isExistAloneP = false;
                    $this->isPStart = false;
                }
            }
            //段落标签
            $ul = '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>';
            //解析ul、ol-li节点
            foreach ($node->childNodes as $liNode) {
                //判断节点名称
                if ($liNode->nodeName != 'li') {
                    continue;
                }
                //段落开始
                $this->body .= $ul;
                $this->isPStart = true;
                if ($liNode->childNodes->length > 0) {
                    foreach ($liNode->childNodes as $childNode) {
                        $this->parseWord($childNode, $styles);
                    }
                }
                //段落结束
                if ($this->isPStart) {
                    $this->body .= Word::WP_END;
                    $this->isPStart = false;
                }
            }
        }
        //不在继续迭代
        $this->continueEach = false;
    }


    /**
     * 获取主体内容
     *
     * @return mixed
     */
    public function getBody()
    {
        return $this->body;
    }


    public function getRelationshipStr()
    {
        return $this->relationshipStr;
    }


    public function getPkgpartStr()
    {
        return $this->pkgpartStr;
    }


    /**
     * Parse style
     *
     * @param \DOMAttr $attribute
     * @param array $styles
     * @return array
     */
    protected static function parseStyle($attribute, $styles)
    {
        $properties = explode(';', trim($attribute->value, " \t\n\r\0\x0B;"));

        foreach ($properties as $property) {
            list($cKey, $cValue) = array_pad(explode(':', $property, 2), 2, null);
            $cValue = trim($cValue);
            switch (trim($cKey)) {
                case 'text-decoration':
                    switch ($cValue) {
                        case 'underline':
                            $styles['underline'] = 'single';
                            break;
                        case 'line-through':
                            $styles['strikethrough'] = true;
                            break;
                    }
                    break;
                case 'text-align':
                    $styles['alignment'] = self::mapAlign($cValue);
                    break;
                case 'display':
                    $styles['hidden'] = $cValue === 'none' || $cValue === 'hidden';
                    break;
                case 'direction':
                    $styles['rtl'] = $cValue === 'rtl';
                    break;
                case 'font-size':
                    $styles['size'] = Converter::cssToPoint($cValue);
                    break;
                case 'font-family':
                    $cValue = array_map('trim', explode(',', $cValue));
                    $styles['name'] = ucwords($cValue[0]);
                    break;
                case 'color':
                    $styles['color'] = trim($cValue, '#');
                    break;
                case 'background-color':
                    $styles['bgColor'] = trim($cValue, '#');
                    break;
                case 'line-height':
                    $matches = array();
                    $line_height = 240;
                    if (preg_match('/([0-9]+\.?[0-9]*[a-z]+)/', $cValue, $matches)) {
                        //matches number with a unit, e.g. 12px, 15pt, 20mm, ...
                        $spacingLineRule = LineSpacingRule::EXACT;
                        $spacing = Converter::cssToTwip($matches[1]);
                    } elseif (preg_match('/([0-9]+)%/', $cValue, $matches)) {
                        //matches percentages
                        $spacingLineRule = LineSpacingRule::AUTO;
                        //we are subtracting 1 line height because the Spacing writer is adding one line
                        $spacing = ((((int) $matches[1]) / 100) * $line_height) - $line_height;
                    } else {
                        //any other, wich is a multiplier. E.g. 1.2
                        $spacingLineRule = LineSpacingRule::AUTO;
                        //we are subtracting 1 line height because the Spacing writer is adding one line
                        $spacing = 0;
                    }
                    $styles['spacingLineRule'] = $spacingLineRule;
                    $styles['line-spacing'] = $spacing;
                    break;
                case 'letter-spacing':
                    $styles['letter-spacing'] = Converter::cssToTwip($cValue);
                    break;
                case 'text-indent':
                    $styles['indentation']['firstLine'] = Converter::cssToTwip($cValue);
                    break;
                case 'font-weight':
                    $tValue = false;
                    if (preg_match('#bold#', $cValue)) {
                        $tValue = true; // also match bolder
                    }
                    $styles['bold'] = $tValue;
                    break;
                case 'font-style':
                    $tValue = false;
                    if (preg_match('#(?:italic|oblique)#', $cValue)) {
                        $tValue = true;
                    }
                    $styles['italic'] = $tValue;
                    break;
                case 'margin-top':
                    $styles['spaceBefore'] = Converter::cssToPoint($cValue);
                    break;
                case 'margin-bottom':
                    $styles['spaceAfter'] = Converter::cssToPoint($cValue);
                    break;
                case 'margin-left':
                    $styles['spaceLeft'] = Converter::cssToPoint($cValue);
                    break;
                case 'border-color':
                    self::mapBorderColor($styles, $cValue);
                    break;
                case 'border-width':
                    $styles['borderSize'] = Converter::cssToPoint($cValue);
                    break;
                case 'border-style':
                    $styles['borderStyle'] = self::mapBorderStyle($cValue);
                    break;
                case 'width':
                    if (preg_match('/([0-9]+[a-z]+)/', $cValue, $matches)) {
                        $styles['width'] = Converter::cssToTwip($matches[1]);
                        $styles['unit'] = TblWidth::TWIP;
                    } elseif (preg_match('/([0-9]+)%/', $cValue, $matches)) {
                        $styles['width'] = $matches[1] * 50;
                        $styles['unit'] = TblWidth::PERCENT;
                    } elseif (preg_match('/([0-9]+)/', $cValue, $matches)) {
                        $styles['width'] = $matches[1];
                        $styles['unit'] = TblWidth::AUTO;
                    }
                    break;
                case 'border':
                    if (preg_match('/([0-9]+[^0-9]*)\s+(\#[a-fA-F0-9]+)\s+([a-z]+)/', $cValue, $matches)) {
                        $styles['borderSize'] = Converter::cssToPoint($matches[1]);
                        $styles['borderColor'] = trim($matches[2], '#');
                        $styles['borderStyle'] = self::mapBorderStyle($matches[3]);
                    }
                    break;
            }
        }

        return $styles;
    }



    /**
     * 解析class
     *
     * @param $attribute
     * @param $styles
     * @return mixed
     */
    protected static function parseClass($attribute, $styles)
    {
        $properties = explode(';', trim($attribute->value, " \t\n\r\0\x0B;"));

        foreach ($properties as $property) {
            list($cKey, $cValue) = array_pad(explode(':', $property, 2), 2, null);
            //$cValue = trim($cValue);
            switch (trim($cKey)) {
                case 'text-decoration':
                    $styles['underline'] = Jc::SINGLE;
                    break;
                case 'mathjye-underpoint':
                    $styles['italic'] = Jc::HOT;
                    break;
                case 'mathjye-underwave':
                    $styles['underline'] = Jc::WAVE;
                    break;
                case 'mathjye-alignright':
                    $styles['alignment'] = Jc::RIGHT;
                    break;
                case 'mathjye-alignleft':
                    $styles['alignment'] = Jc::LEFT;
                    break;
                case 'mathjye-aligncenter':
                    $styles['alignment'] = Jc::CENTER;
                    break;
            }
        }

        return $styles;
    }



    /**
     * 解析图片
     *
     * @param \DOMNode $node
     * @param array $styles
     * @return mixed
     * @throws \Exception
     */
    protected function parseImage($node, $styles)
    {
        $style = array('width' => '', 'height' => '');
        $src = null;
        foreach ($node->attributes as $attribute) {
            switch ($attribute->name) {
                case 'src':
                    $src = $attribute->value;
                    break;
                case 'width':
                    $width = $attribute->value;
                    $style['width'] = $width;
                    $style['unit'] = Style::UNIT_PX;
                    break;
                case 'height':
                    $height = $attribute->value;
                    $style['height'] = $height;
                    $style['unit'] = Style::UNIT_PX;
                    break;
                case 'style':
                    $styleattr = explode(';', $attribute->value);
                    foreach ($styleattr as $attr) {
                        if (strpos($attr, ':')) {
                            list($k, $v) = explode(':', $attr);
                            $k = trim($k);
                            $v = trim($v);
                            switch ($k) {
                                case 'float':
                                case 'FLOAT':
                                    if ($v == 'right') {
                                        $style['hPos'] = Style::POS_RIGHT;
                                        $style['hPosRelTo'] = Style::POS_RELTO_PAGE;
                                        $style['pos'] = Style::POS_RELATIVE;
                                        $style['wrap'] = Style::WRAP_TIGHT;
                                        $style['overlap'] = true;
                                    }
                                    if ($v == 'left') {
                                        $style['hPos'] = Style::POS_LEFT;
                                        $style['hPosRelTo'] = Style::POS_RELTO_PAGE;
                                        $style['pos'] = Style::POS_RELATIVE;
                                        $style['wrap'] = Style::WRAP_TIGHT;
                                        $style['overlap'] = true;
                                    }
                                    break;
                                case 'width':
                                    $style['width'] = $v;
                                    $style['unit'] = Style::UNIT_PX;
                                    break;
                                case 'height':
                                    $style['height'] = $v;
                                    $style['unit'] = Style::UNIT_PX;
                                    break;
                            }
                        }
                    }
                    break;
            }
        }

        $id = uniqid();
        //base64
        $base64Image = '';
        //图片后缀
        $suffix = 'png';
        if (strpos($src, 'data:image') !== false) {

            $match = array();
            preg_match('/data:image\/(\w+);base64,(.+)/', $src, $match);
            $base64Image = $match[2];
            $suffix = $match[1];
            if (empty($style['width'])) {
                $imageData = $this->getImageSize(base64_decode($match[2]));
                if (is_array($imageData)) {
                    list($actualWidth, $actualHeight, $imageType) = $imageData;
                    $style['width'] = $actualWidth . Style::UNIT_PX;
                    $style['height'] = $actualHeight . Style::UNIT_PX;
                }
            }
        }
        $src = urldecode($src);

        if (!is_file($src)) {
            if ($imgBlob = self::sendRequest($src)) {

                $match = array();
                preg_match('/.+\.(\w+)$/', $src, $match);
                $suffix = $match[1];
                $base64Image = base64_encode($imgBlob);
                if (empty($style['width'])) {
                    $imageData = $this->getImageSize($imgBlob);
                    if (is_array($imageData)) {
                        list($actualWidth, $actualHeight, $imageType) = $imageData;
                        $style['width'] = $actualWidth . Style::UNIT_PX;
                        $style['height'] = $actualHeight . Style::UNIT_PX;
                    }
                }
            }
        }
        if ($base64Image) {
            //创建段落
            $this->createParagraph($styles);
            //relationship信息
            $image = 'image' . $id . $suffix;
            $data = [$id, $image];
            $this->relationshipStr[] = self::replaceStr(Word::RELATIONSHIP, $data);
            //正文图片信息
            //图片样式
            $imageStyle = [
                'width:' . $style['width'],
                'height:' . $style['height']
            ];
            if (isset($style['hPos'])) {
                $imageStyle = array_merge($imageStyle, [
                    'margin-left:0pt',
                    'margin-top:0pt',
                    'position:' . $style['pos'],
                    'mso-position-horizontal:' . $style['hPos'],
                    'mso-position-vertical:top',
                    'mso-position-horizontal-relative:' . $style['hPosRelTo'],
                    'mso-position-vertical-relative:line'
                ]);
            }
            $this->body .= self::replaceStr(Word::DOCUMENT_PIC_MARK, [$id, implode(';', $imageStyle), $id]);
            //pkgpart信息
            $this->pkgpartStr[] = self::replaceStr(Word::PKG_PART, [$image, $suffix, $base64Image]);
        }
    }




    private function getImageSize($source)
    {
        if (!function_exists('getimagesizefromstring')) {
            $uri = 'data://application/octet-stream;base64,' . base64_encode($source);
            $result = @getimagesize($uri);
        } else {
            $result = @getimagesizefromstring($source);
        }

        return $result;
    }



    private static function replaceStr($str, $newArr)
    {
        preg_match_all('/#(.*?)#/is', $str, $match);
        foreach ($match[0] as $k => $item) {
            $str = str_replace($item, $newArr[$k], $str);
        }
        return $str;
    }


    /**
     * Transforms a CSS border style into a word border style
     *
     * @param string $cssBorderStyle
     * @return null|string
     */
    protected static function mapBorderStyle($cssBorderStyle)
    {
        switch ($cssBorderStyle) {
            case 'none':
            case 'dashed':
            case 'dotted':
            case 'double':
                return $cssBorderStyle;
            default:
                return 'single';
        }
    }

    protected static function mapBorderColor(&$styles, $cssBorderColor)
    {
        $numColors = substr_count($cssBorderColor, '#');
        if ($numColors === 1) {
            $styles['borderColor'] = trim($cssBorderColor, '#');
        } elseif ($numColors > 1) {
            $colors = explode(' ', $cssBorderColor);
            $borders = array('borderTopColor', 'borderRightColor', 'borderBottomColor', 'borderLeftColor');
            for ($i = 0; $i < min(4, $numColors, count($colors)); $i++) {
                $styles[$borders[$i]] = trim($colors[$i], '#');
            }
        }
    }


    /**
     * 对齐样式
     *
     * @param string $cssAlignment
     * @return string|null
     */
    protected static function mapAlign($cssAlignment)
    {
        switch ($cssAlignment) {
            case 'right':
                return Jc::END;
            case 'center':
                return Jc::CENTER;
            case 'justify':
                return Jc::BOTH;
            default:
                return Jc::START;
        }
    }


    /**
     * Parse line break
     *
     * @param array &$styles
     */
    protected function parseLineBreak(&$styles)
    {
        //alignment applies on paragraph, not on font. Let's copy it there
        if (isset($styles['font']['alignment']) && is_array($styles['paragraph'])) {
            $styles['paragraph']['alignment'] = $styles['font']['alignment'];
        }
        //创建段落
        $this->createParagraph($styles);
        $this->body .= Word::BR;
    }


    /**
     * Parse link node
     *
     * @param \DOMNode $node
     * @param array $styles
     */
    protected function parseLink($node, &$styles)
    {
        $target = null;
        foreach ($node->attributes as $attribute) {
            switch ($attribute->name) {
                case 'href':
                    $target = $attribute->value;
                    break;
            }
        }
        $styles['font'] = self::parseInlineStyle($node, $styles['font']);
    }



    /**
     * CURL发送Request请求,含POST和REQUEST
     * @param string $url     请求的链接
     * @param mixed  $params  传递的参数
     * @param string $method  请求的方法
     * @param mixed  $options CURL的参数
     * @return array
     */
    public static function sendRequest($url, $params = [], $method = 'GET', $options = [])
    {
        $method = strtoupper($method);
        $protocol = substr($url, 0, 5);
        $query_string = is_array($params) ? http_build_query($params) : $params;
        $ch = curl_init();
        $defaults = [];
        if ('GET' == $method) {
            $geturl = $query_string ? $url . (stripos($url, "?") !== false ? "&" : "?") . $query_string : $url;
            $defaults[CURLOPT_URL] = $geturl;
        } else {
            $defaults[CURLOPT_URL] = $url;
            if ($method == 'POST') {
                $defaults[CURLOPT_POST] = 1;
            } else {
                $defaults[CURLOPT_CUSTOMREQUEST] = $method;
            }
            $defaults[CURLOPT_POSTFIELDS] = $params;
        }

        $defaults[CURLOPT_HEADER] = false;
        $defaults[CURLOPT_USERAGENT] = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.98 Safari/537.36";
        $defaults[CURLOPT_FOLLOWLOCATION] = true;
        $defaults[CURLOPT_RETURNTRANSFER] = true;
        //$defaults[CURLOPT_CONNECTTIMEOUT] = 3;
        //$defaults[CURLOPT_TIMEOUT] = 3;

        // disable 100-continue

        curl_setopt($ch, CURLOPT_HTTPHEADER, array('Expect:'));


        if ('https' == $protocol) {
            $defaults[CURLOPT_SSL_VERIFYPEER] = false;
            $defaults[CURLOPT_SSL_VERIFYHOST] = false;
        }
        curl_setopt_array($ch, (array)$options + $defaults);


        $ret = curl_exec($ch);
        curl_close($ch);
        return $ret;
    }



}