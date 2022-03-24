# html to word

#### 介绍
html标签文本转word标签文本
支持mathml解析

## 安装

Package available on [Composer](https://packagist.org/packages/zmh/html-parser-word).

If you're using Composer to manage dependencies, you can use

    $ composer require zmh/html-parser-word

#### 使用说明

```
$html = 'hello <b>world</b>!';
$parse = new Parse();
$parse->doHtml2Word($html);
echo $parse->getBody();
```