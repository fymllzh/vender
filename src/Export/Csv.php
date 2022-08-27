<?php
/*
 * @Description: csv导出
 * @version: 1.0
 * @Author: zhihao.li
 * @Date: 2022-1-13 20:44:17
 */

namespace FFMBase\Export;

class Csv extends AbstractExport
{
    protected $ext = '.csv';
    protected $fileMap = [];

    /**
     * @desc csv追加导出多文件并支持大文件拆分
     * @param $data
     * @return mixed|void
     */
    public function export($data)
    {
        //入参校验
        $params = $this->validatorParams($data);
        $name   = $params['name'];

        //创建本地文件
        if (!$this->append) {
            if (!is_dir($this->params['upload_temp_dir'] . $this->random) && !mkdir($this->params['upload_temp_dir'] . $this->random, 0777, true)) {
                throw new \Exception('folder does not exist or failed to create');
            }
        }

        $csvBaseFile = $this->params['upload_temp_dir'] . $this->random . '/' . $name;
        $filename    = $csvBaseFile . ((isset($this->fileMap[$name]['suffix_num']) && $this->fileMap[$name]['suffix_num']) ? '_' . $this->fileMap[$name]['suffix_num'] : '') . $this->ext;
        $fp          = $this->openFile($filename);

        //首次则写入title
        if (!$this->append) {
            $this->writeTitle($fp,$params['title']);
            //初始化fileMap
            $this->fileMap[$name]['row_num'] = 0;
            $this->fileMap[$name]['suffix_num'] = (isset($this->fileMap[$name]['suffix_num']) && $this->fileMap[$name]['suffix_num']) ? $this->fileMap[$name]['suffix_num'] : 0;
        }

        foreach ($params['body'] as $key => $row) {
            $this->fileMap[$name]['row_num']++;
            $tmp = [];

            //key-value格式写入数据
            foreach ($params['data_fields'] as $field) {
                if (isset($row[$field])) {
                    $tmp[] = (string)$row[$field];
                } else {
                    if ('excelIndex' == $field) {
                        $tmp[] = $this->fileMap[$name]['row_num'];
                    }
                }
            }
            fputcsv($fp, $tmp, ",", '"');

            //分文件
            if ($this->fileMap[$name]['row_num'] == $this->maxNumber) {
                $this->fileMap[$name]['suffix_num']++;
                //临界点处理
                if(isset($params['body'][$key + 1])) {
                    //关掉上一个csv
                    $this->closeFile($fp);

                    //兼容首次就需要分文件的场景
                    if (!$this->append) {
                        $this->tmpFiles[] = $filename;
                    }

                    //生成下一个csv并写入title
                    $filename = $csvBaseFile . '_' . $this->fileMap[$name]['suffix_num'] . $this->ext;
                    $fp       = $this->openFile($filename);
                    $this->writeTitle($fp,$params['title']);

                    //重置参数
                    $this->append = false;
                    $this->fileMap[$name]['row_num'] = 0;
                } else {
                    $this->append = null;
                }
            }
        }

        //关闭文件
        $this->closeFile($fp);

        //文件写入数组
        if (!$this->append) {
            $this->tmpFiles[] = $filename;
        }
    }

    /**
     * @desc 打开文件
     * @param $filename
     * @return resource
     */
    private function openFile($filename)
    {
        $fp = fopen($filename, 'a+');
        if (!$fp) {
            throw new \Exception('failed to open temporary file');
        }

        return $fp;
    }

    /**
     * @desc 写入title
     * @param $fp
     * @param $title
     */
    private function writeTitle($fp,$title)
    {
        fwrite($fp, chr(0xEF) . chr(0xBB) . chr(0xBF));
        $ret = fputcsv($fp, $title, ",", '"');
        if (false === $ret) {
            fclose($fp);
            throw new \Exception('failed to write header');
        }
    }

    /**
     * @desc 关闭文件
     * @param $fp
     */
    private function closeFile($fp)
    {
        $res = fclose($fp);
        if (false === $res) {
            throw new \Exception('failed to close temporary file');
        }
    }
}