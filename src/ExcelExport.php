<?php

namespace Yxg\Excel;

/**
 * @title 导出excel
 *
 * @desc ExcelExport.php.php
 *
 *
 * @create_time 2022-03-08 15:23
 */
class ExcelExport
{
    /**
     * 测试方法.
     */
    public function test()
    {
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

        return $this->export('统计导出', $cell, $data, $merge_key, $merge_columns);
    }

    /**
     * 导出
     * 参考 test方法，或 https://github.com/mk-j/PHP_XLSXWriter/tree/master/examples.
     *
     * @param string $filename     文件名称
     * @param array  $cell_data    标题头
     * @param array  $row_data     数据
     * @param string $merge_key    需要合并的唯一键
     * @param array  $merge_column 需要合并的数据字段，数组格式
     * @param string $path 保存路径
     *
     * @throws Exception
     */
    public function export(string $filename = '导出列表', array $cell_data = [], array $row_data = [], string $merge_key = '', array $merge_column = [], $file_type = 'xlsx', $path = '/')
    {
        $writer = new XLSXWriter();
        $writer->setTitle('科技有限公司');
        // $writer->setAuthor(request()->param('userName'));
        $writer->setAuthor('userName');
        $writer->setCompany('科技有限公司');
        $style = ['border' => 'left,right,top,bottom', 'border-style' => 'thin', 'valign' => 'center', 'wrap_text' => 'true', 'height' => 20];
        $hds = [];
        $this->config2HDS($cell_data, $hds);
        // 获取表格标题的每一行分布情况
        $hd = $this->getHdsRowDs($hds);
        // 获取每一列的宽度
        $hw = $this->getHdsWidths($hds);
        // 获取每一列对应的field
        $hf = $this->getHdsRowFs($hds);
        // 获取标题合并的坐标信息
        $hm = $this->getHdsMergeInfo($hds);
        // 获取行数据合并的坐标信息
        $row_hms = $this->getMergeRowInfo($row_data, $merge_key, $merge_column, count($hd), $hf);
        foreach ($row_hms as $row_hm) {
            array_push($hm, $row_hm);
        }
        $hs = [
            'suppress_row' => true,
            'widths' => $hw,
        ];
        $writer->writeSheetHeader('Sheet1', array_pad([], count($hw), 'string'), $col_options = $hs);

        foreach ($hd as $dhi => $dh) {
            $writer->writeSheetRow('Sheet1', $dh, array_merge($style, $dhi == 0 ? ['halign' => 'center', 'font-style' => 'bold'] : ['halign' => 'center']));
        }
        $data = [];
        $i = 0;
        foreach ($row_data as $d) {
            $line = [++$i];
            foreach ($hf as $hk => $v) {
                $line[$hk] = $d[$v];
            }
            $data[] = $line;
        }
        foreach ($data as $d) {
            $writer->writeSheetRow('Sheet1', $d, $style);
        }
        // 执行合并单元格
        foreach ($hm as $mg) {
            $writer->markMergedCell('Sheet1', $mg[0], $mg[1], $mg[2], $mg[3]);
        }

        // 目录不存在，则创建
        if (!file_exists($path)) {
            // iconv防止中文名乱码
            mkdir(iconv('UTF-8', 'GBK', $path), 0777, true);
        }
        $writer->writeToFile($path.$filename.'.'.$file_type);

        return $path.$filename.'.'.$file_type;
    }

    /**
     * 获取需要合并的行的坐标
     * 坐标：行起，列起，行止，列止.
     *
     * @param array  $data          数据
     * @param string $merge_key     唯一键
     * @param array  $merge_columns 需要合并的数组
     * @param int    $exist_row_num 标题行的个数
     * @param array  $header_fields 标题的字段分布
     */
    private function getMergeRowInfo($data, $merge_key, $merge_columns, $exist_row_num, $header_fields): array
    {
        // 计算根据唯一键找找到对应的重复次数
        $uni_array = array_filter(array_count_values(array_column($data, $merge_key)), function ($value) {
            return $value > 1;
        });
        $temp = [];
        foreach ($data as $key => $value) {
            foreach ($uni_array as $uni_key => $count) {
                if ($value[$merge_key] == $uni_key) {
                    foreach ($header_fields as $f_key => $field) {
                        if (in_array($field, $merge_columns)) {
                            $start_row = $key;
                            $end_row = $start_row + $count - 1;
                            // 坐标：行起，列起，行止，列止
                            $temp[] = [$start_row + $exist_row_num, $f_key, $end_row + $exist_row_num, $f_key];
                        }
                    }
                    unset($uni_array[$uni_key]);
                }
            }
        }

        return $temp;
    }

    /**
     * 返回数组的维度.
     *
     * @param  [type] $arr [description]
     *
     * @return [type]      [description]
     */
    private function arrayLevel($arr)
    {
        $al = [0];
        function aL($arr, &$al, $level = 0)
        {
            if (is_array($arr)) {
                ++$level;
                $al[] = $level;
                foreach ($arr as $v) {
                    aL($v, $al, $level);
                }
            }
        }

        aL($arr, $al);

        return max($al);
    }

    /**
     * 获取每一列的宽度.
     *
     * @return array
     */
    private function getHdsWidths($hds)
    {
        $re = [];
        foreach ($hds as $hid => $h) {
            foreach ($h as $lid => $l) {
                if (array_key_exists(1, $l) && (int) $l[1] > 0) {
                    $re[$lid] = (int) $l[1];
                }
            }
        }

        return $re;
    }

    /**
     * 获取每一列对应的field.
     *
     * @return array
     */
    private function getHdsRowFs($hds)
    {
        $re = [];
        foreach ($hds as $hid => $h) {
            foreach ($h as $lid => $l) {
                if (array_key_exists(2, $l)) {
                    $re[$lid] = $l[2];
                }
            }
        }

        return $re;
    }

    /**
     * 获取表格标题的每一行分布情况.
     *
     * @return array
     */
    private function getHdsRowDs($hds)
    {
        $re = [];
        $lm = 0;
        foreach ($hds as $hid => $h) {
            $ls = max(array_keys($h));
            if ($ls > $lm) {
                $lm = $ls;
            }
        }
        foreach ($hds as $hid => $h) {
            $re[$hid] = array_pad([], $lm, '');
            foreach ($h as $lid => $l) {
                $re[$hid][$lid] = $l[0];
            }
        }

        return $re;
    }

    /**
     * 获取合并的坐标信息.
     *
     * @return array
     */
    private function getHdsMergeInfo($hds)
    {
        $re = [];
        foreach ($hds as $hid => $h) {
            foreach ($h as $lid => $l) {
                if (array_key_exists('merge', $l)) {
                    $re[] = $l['merge'];
                } else {
                    if (!empty($l[0]) && $hid < (count($hds) - 1) && (isset($hds[$hid + 1][$lid][0]) && $hds[$hid + 1][$lid][0] == '')) {
                        $re[] = [$hid, $lid, $hid + 1, $lid];
                    }
                }
            }
        }

        return $re;
    }

    /**
     * 组装表头个数.
     */
    private function config2HDS($config, &$hds, int $h = 0, int $l = 0): int
    {
        $hs = [];
        $i = -1;
        foreach ($config as $v) {
            ++$i;
            if (!array_key_exists($h, $hds)) {
                $hds[$h] = [];
            }

            if ($v[0] == '-') {
                $w = $this->config2HDS($v[2], $hds, $h + 1, $l + $i);
                $hds[$h][$l + $i] = [$v[1], 'merge' => [$h, $l + $i, $h, $l + $i + $w]];
                for ($k = 1; $k <= $w; ++$k) {
                    $hds[$h][$l + $i + $k] = [''];
                }
                $i += $w;
            } else {
                $hds[$h][$l + $i] = [$v[1], $v[2], $v[0]];
            }
        }

        return $i;
    }
}
