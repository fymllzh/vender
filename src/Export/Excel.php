<?php
/*
 * @Description: Exce导出
 * @version: 1.0
 * @Author: panjia
 * @Date: 2022-1-13 20:44:17
 */

namespace FFMBase\Export;

class Excel extends AbstractExport
{
    protected $ext = '.xlsx';
    protected $sheetIndex;
    private $lastRowIndex;

    public function __construct($params)
    {
        parent:: __construct($params);

        ini_set('memory_limit', '2048M');
        $this->objPHPExcel = new \PHPExcel();
        $this->filename    = $this->getFilename(); //文件绝对路径
        $this->suffix      = '';    // 后缀名字
        $this->suffix_num  = 0;     // 后缀数量
        $this->row_num     = 0;     // 临时数量
        $this->next        = 0;     // 下一个sheet的延续数量
        $this->nextTmp     = [];    // 下一个sheet的延续数据
        $this->flag        = false; // sheet拆分标识
    }

    /**
     * @param $data
     *
     * @return mixed|void
     * @throws \PHPExcel_Exception
     */
    public function export($data)
    {
        // 验证
        $this->validatorParams($data);

        // 计算每次写入的数量
        $count = count($data['body']);

        // 累计每个sheet 的数量
        $this->row_num += $count;

        // 超过 指定数量时 做拆分
        if ($this->row_num >= $this->maxNumber) {
            $this->next = $this->row_num - $this->maxNumber;
            ++$this->next; // 解决临界点问题
            $this->row_num = 0;
        }

        // 重新生成下一个sheet
        if ($this->flag) {
            $this->append = false;
            $this->sheetIndex++;
            $this->suffix_num++;
            $this->suffix = $this->suffix_num ? '_'.$this->suffix_num : '';
        }
        $this->flag = false;

        // 创建sheet
        if ($this->sheetIndex != 0 && $this->append == false) {
            $this->objPHPExcel->createSheet();
        }

        // sheet页码
        if ($this->sheetIndex === null) {
            $this->sheetIndex = 0;
        }

        // 设置sheet页码
        $this->objPHPExcel->setActiveSheetIndex($this->sheetIndex);

        //数据
        $sheet = $data['sheet'].$this->suffix;//sheet 名称
        $head = $data['head']; //表头信息
        $body = $data['body']; //内容

        //初始化表格行列标记
        $row = 1;
        $column = 'A';
        $size = 14;
        $excel = self::excel($head, $column);
        $dataFields = $excel['value'];//字段
        $label = $excel['label'];//表头
        $maxColumn = $excel['maxColumn'];//最大excel列

        //获得当前sheet
        $curSheet = $this->objPHPExcel->getActiveSheet();
        if (!$this->append) {
            $curSheet->setTitle($sheet);
        }

        $column_row = $column . $row . ':' . $maxColumn . $row;

        //设置表头样式
        $curSheet->getStyle($column_row)->getFill()
            ->setFillType(\PHPExcel_Style_Fill::FILL_SOLID)
            ->getStartColor()->setRGB('EFF3F6');


        $curSheet->getRowDimension($row)->setRowHeight(24);
        $curSheet->getStyle($column_row)->getFont()->setSize($size);
        $curSheet->getDefaultStyle()->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT);//水平
        $curSheet->getDefaultStyle()->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);//竖直
        $curSheet->getStyle($column_row)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);//
        $curSheet->getDefaultColumnDimension()->setWidth(15);
        $curSheet->freezePane("A2");
        //$curSheet->getDefaultStyle()->getAlignment()->setWrapText(true);//
        //数据组装并写入sheet
        //$lastRowIndex = $this->lastRowIndex;
        if (empty($this->lastRowIndex[$sheet])) {
            $this->lastRowIndex[$sheet] = 1;
        }
        $rows = [];
        $startCell = 'A1';
        if ($this->append) {
            $startCell = 'A' . ($this->lastRowIndex[$sheet] + 1);
        } else {
            $rows[] = $label;
        }

        // 把上次遗留的数据压进本次sheet里
        $this->nextTmp = array_reverse($this->nextTmp);
        foreach ($this->nextTmp as $v) {
            array_unshift($body, $v);
        }
        unset($v);

        $this->nextTmp = [];

        // 多余数据切割
        if ($this->next > 0) {
            --$this->next; // 清除临界点临时数量

            // 真正剩余数量
            if ($this->next > 0) {
                $this->nextTmp = array_slice($body, '-'.$this->next);

                // 拿到差集
                $body = array_slice($body, 0, count($body)-$this->next);

                // 剩余数累加到下一次循环
                $this->row_num += count($this->nextTmp);
            }
            $this->flag = true;
        }
        // excel 数据组建
        foreach ($body as $row) {
            $tmp = [];
            foreach ($dataFields as $field) {
                if (isset($row[$field])) {
                    $tmp[] = (string)$row[$field];
                } else {
                    $defaultv = "";
                    if ('excelIndex' == $field) {
                        $defaultv = $this->lastRowIndex[$sheet];
                    }
                    $tmp[] = $defaultv;
                }
            }
            $rows[$this->lastRowIndex[$sheet]] = $tmp;
            $this->lastRowIndex[$sheet]++;
        }

        // 推送给excel
        $this->fromArray($curSheet, $rows, null, $startCell);

        unset($rows);
        unset($startCell);

        // 递归处理剩余数据
        if (!empty($this->nextTmp)) {
            $nextTmp['sheet'] = $data['sheet'];
            $nextTmp['head'] = $data['head'];
            $nextTmp['body'] = $this->nextTmp;
            $this->nextTmp = [];
            $this->next = 0;
            $this->row_num = 0;
            $this->export($nextTmp);
            unset($nextTmp);
        }

    }

    /**
     * 返回Excel所需键值对。。
     * @param $data [[加偏移列最大支持702位]]
     * @param string $offset 偏移列
     * @return array
     */
    public static function excel($data, $offset = 'A')
    {
        $alphabet = str_split("ABCDEFGHIJKLMNOPQRSTUVWXYZ");
        $excel = [
            "label" => [],
            "value" => [],
            "maxColumn" => "A"
        ];
        ##获取偏移量。。
        $i = self::abcTo123($offset);
        foreach ($data as $k => $v) {
            $f = floor($i / 26);
            $m = $i % 26;
            $e = $f ? $alphabet[$f - 1] . $alphabet[$m] : $alphabet[$m];
            $excel['label'][$e] = $k;
            $excel['value'][$e] = $v;
            $excel['maxColumn'] = $e;
            $i++;
        }
        return $excel;
    }


    /**
     * 把Excel所需的列字母转为数字。。
     * @param string $abc
     * @return false|float|int|string
     */
    public static function abcTo123($abc = "A")
    {
        $alphabet = str_split("ABCDEFGHIJKLMNOPQRSTUVWXYZ");
        $offsetList = array_reverse(str_split($abc));
        $i = 0;
        foreach ($offsetList as $k => $v) {
            $key = array_search($v, $alphabet);
            $i += $k ? ($key + 1) * 26 : $key;
        }
        return $i;
    }

    /**
     * @desc 生成文件路径
     * @return string
     */
    private function getFilename()
    {
        $filename = $this->params['upload_temp_dir'] . $this->filename() . $this->ext;
        return $filename;
    }

    /**
     * 生成文件名。。
     * @return string
     * @throws \Exception
     */
    private function filename()
    {
        return md5($this->randomString(9) . mt_rand(1000, 9999)) . md5(date('YmdHis') . mt_rand(1000, 9999));
    }

    /**
     * 返回随机字符串。。
     * @param $len
     * @return string
     * @throws \Exception
     */
    private function randomString($len)
    {
        $charPool = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
        $charPoolSize = strlen($charPool);
        $string = "";
        while ($len) {
            $string .= $charPool[random_int(0, $charPoolSize - 1)];
            --$len;
        }
        return $string;
    }

    /**
     * 替换掉 composer 中的 fromArray 方法
     * @param \PHPExcel_Worksheet $worksheet
     * @param null $source
     * @param null $nullValue
     * @param string $startCell
     *
     * @return \PHPExcel_Worksheet
     * @throws \PHPExcel_Exception
     */
    private function fromArray(\PHPExcel_Worksheet $worksheet, $source = null, $nullValue = null, $startCell = 'A1')
    {
        if (is_array($source)) {
            //    Convert a 1-D array to 2-D (for ease of looping)
            if (!is_array(end($source))) {
                $source = array($source);
            }

            // start coordinate
            list ($startColumn, $startRow) = \PHPExcel_Cell::coordinateFromString($startCell);

            // Loop through $source
            foreach ($source as $rowData) {
                $currentColumn = $startColumn;
                foreach ($rowData as $cellValue) {
                    if ($cellValue != $nullValue) {
                        // Set cell value
                        $worksheet->getCell($currentColumn . $startRow)->setValueExplicit($cellValue);
                    }
                    ++$currentColumn;
                }
                ++$startRow;
            }
        } else {
            throw new \PHPExcel_Exception("Parameter \$source should be an array.");
        }
        return $worksheet;
    }
}