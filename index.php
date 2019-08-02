<?php

use core\exportAction;
require 'vendor/autoload.php';
spl_autoload_register(function ($class) {
    $class = str_replace('\\', '/', $class);
    $filePath = __DIR__ . DIRECTORY_SEPARATOR . $class . '.php';
    if (file_exists($filePath)) {
        include $filePath;
    }
});
$start = time();
function createData(int $number)
{
    $big = [];
    $dic = "最想要看到是你的微笑在我的眼中你是最好肉麻的调调你不会知道我爱的静悄悄我该怎么往下聊全都怪我太胆小只会看着你傻笑怎么办才好可我真的没有想到你把我拥入怀抱世界突然变得好安静只剩心跳的声音坚定了我爱你的决心此刻你就是唯一世界突然变得好安静不敢用力的呼吸因为我害怕这是梦境不小心会惊醒最想要看到是你的微笑在我的眼中你是最好肉麻的调调你不会知道我爱的静悄悄我该怎么往下聊全都怪我太胆小只会看着你傻笑怎么办才好可我真的没有想到你把我拥入怀抱世界突然变得好安静只剩心跳的声音坚定了我爱你的决心此刻你就是唯一世界突然变得好安静不敢用力的呼吸因为我害怕这是梦境不小心会惊醒世界突然变得好安静只剩心跳的声音坚定了我爱你的决心此刻你就是唯一世界突然变得好安静不敢用力的呼吸因为我害怕这是梦境不小心会惊醒吉他陈羿淳和声高漠伊王中易和声编写高漠伊王中易混音李琰祥制作人张雯杰";
    for ($i = 1; $i <= $number; $i++) {
        $tLen = mb_strlen($dic, 'UTF-8');
        $start = rand(0, $tLen);
        $len = rand(2, 4);
        $name = mb_substr($dic, $start, $len, 'UTF-8');
        $big[] = [
            'id' => $i,
            'name' => $name,
            'age' => rand(10, 80),
            'mobile' => '15542654821',
            'address' => '四川省成都市武侯区桐梓林凯莱帝景B座6A',
        ];
    }
    return $big;
}
$number =15000;
$data = createData($number);
$consume =time() - $start;
echo  "创建数据消耗时间：{$consume}\r\n";
$options['data'] = $data;

/**
 * 数据对应的栏目名称
 */
$columnNames = ['主键', '姓名', '年龄', '手机', '地址'];
$options['columnNames'] = $columnNames;

/**
 * 导出操作
 */
$exportAction = new exportAction($options);

$exportAction->setTitle("{$number}条数据测试");
if ($exportAction->export() == false) {
    echo $exportAction->getError();
}


