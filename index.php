<?php
require "vendor/autoload.php";

use Yxg\Excel\ExcelExport;

$excel = new ExcelExport();

// 表头，支出复杂表头自动合并，三个参，第一个为字段field,第二个字段名，第三个参是列宽
$cell = [['-', '统计导出', [
    ['id', '序号', 8],
    ['code', '部门', 16],
    ['name', '姓名', 12],
    ['num', '应出勤', 8],
    ['num', '实际出勤', 10],
    ['num', '实际打卡', 10],
    ['num', '出差天数', 10],
    ['num', '计薪天数', 10],
    ['num', '周末加班', 10],
    ['num', '节日加班', 10],
    ['-', '请休假', [
        ['num', '年假', 6],
        ['num', '婚假', 6],
        ['num', '陪产假', 8],
        ['num', '丧假', 6],
        ['num', '产假', 6],
        ['num', '工伤假', 8],
        ['num', '事假', 6],
        ['num', '病假', 6],
    ]],
    ['-', '夜值', [
        ['num', '夜值A', 8],
        ['num', '夜值B', 8],
    ]],
]]];
$data = [
    ['code' => '001', 'name' => 'Jason1', 'num' => rand(1, 100), 'id' => rand(1, 100)],
    ['code' => '001', 'name' => 'Jason1', 'num' => rand(1, 100), 'id' => rand(1, 100)],
    ['code' => '001', 'name' => 'Jason1', 'num' => rand(1, 100), 'id' => rand(1, 100)],
    ['code' => '002', 'name' => 'Jason2', 'num' => rand(1, 100), 'id' => rand(1, 100)],
    ['code' => '002', 'name' => 'Jason2', 'num' => rand(1, 100), 'id' => rand(1, 100)],
    ['code' => '003', 'name' => 'Jason3', 'num' => rand(1, 100), 'id' => rand(1, 100)],
    ['code' => '004', 'name' => 'Jason4', 'num' => rand(1, 100), 'id' => rand(1, 100)],
    ['code' => '005', 'name' => 'Jason5', 'num' => rand(1, 100), 'id' => rand(1, 100)],
];
// 如需合并数据，则传入这两个字段
$merge_key = 'code';
$merge_columns = ['code', 'name'];
$excel->export('统计导出', $cell, $data, $merge_key, $merge_columns, 'xls','./');
