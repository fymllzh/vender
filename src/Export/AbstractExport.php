<?php
/*
 * @Description: 基类
 * @version: 1.0
 * @Author: zhihao.li
 * @Date: 2022-1-13 20:44:17
 */

namespace FFMBase\Export;

use FFMBase\Util\AliyunOss;
use FFMBase\Util\Common;

abstract class AbstractExport
{
    const MAX_NUMBER = 300000;
    protected $objPHPExcel;
    protected $filename;
    protected $params;
    protected $random;
    protected $append;
    protected $maxNumber;
    protected $tmpFiles = [];

    public function __construct($params)
    {
        $this->params = $params; //参数数组..
        $this->random = date('YmdHis') . '_' . mt_rand(10000000, 99999999); //随机数
        $this->maxNumber = self::MAX_NUMBER; //单文件最大记录数
    }

    /**
     * @desc 导出
     * @param $data
     * @return mixed
     */
    abstract public function export($data);

    /**
     * @desc 多文件时重置参数
     */
    public function resetAppend()
    {
        $this->append = null;
        if ($this->ext == '.xlsx') {
            $this->sheetIndex++;
            $this->row_num = 0;
            $this->suffix_num = 0;
            $this->suffix = '';
            $this->flag = false;
        }
    }

    /**
     * @desc 返回上传文件
     * @return string
     */
    public function getFile()
    {
        // 清理内存&&生成文件
        if (!is_null($this->objPHPExcel)) {
            $this->objPHPExcel->setActiveSheetIndex(0);
            $objWriter = \PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel2007');
            $objWriter->save($this->filename);
            $this->tmpFiles[] = $this->filename;
            $this->objPHPExcel->disconnectWorksheets();
        }

        $files = $this->tmpFiles;

        if (empty($files)) {
            throw new \Exception('empty file');
        }

        if (count($files) == 1) {
            $localFile = $files[0];
        } elseif (count($files) > 1 && $this->ext == '.csv') {
            $objName = $this->random . '.zip';
            $localFile = $this->params['upload_temp_dir'] . $this->random . '/' . $objName;
            $command = escapeshellcmd("zip -mqj $localFile " . implode(' ', $files));
            exec($command, $output, $res);
            if ($res !== 0 || !file_exists($localFile)) {
                throw new \Exception('failed to zip file');
            }
        }

        //上传阿里云
        if(true === $this->params['oss']) {
            return $this->fileToOss($localFile);
        }

        return $localFile;
    }

    /**
     * @desc 特殊字符做一个安全处理
     * @param $filename
     * @return string
     */
    public function formatFilename($filename)
    {
        return $filename ? str_replace(['.', '/', ' ', '$', '\\'], '', $filename) : '';
    }

    /**
     * @desc 设置文件最大处理记录数
     * @param $maxNumber
     */
    public function setMaxNumber()
    {
        $this->maxNumber = 1000000;
    }

    /**
     * @desc 入参校验
     * @param $data
     * @return array
     */
    protected function validatorParams($data)
    {
        if ($this->append === null) {
            $this->append = false;
        } else if ($this->append === false) {
            $this->append = true;
        }

        //首次则校验数据
        if(!$this->append) {
            $validator = [
                'sheet' => 'Required|Str|StrLenGe:1',
                'head'  => 'Required|Obj',
                'body'  => 'Required|Arr',
            ];
            (new Common)::validator($data, $validator);
        }

        $title      = array_keys($data['head']);
        $dataFields = array_values($data['head']);
        $body       = $data['body'];
        $name       = $this->formatFilename($data['sheet']);

        return ['title' => $title, 'body' => $body, 'name' => $name, 'data_fields' => $dataFields];
    }

    /**
     * @desc 上传oss
     * @param $localFile
     * @return string
     */
    public function fileToOss($localFile)
    {
        $s = new AliyunOss($this->params['endpoint'],$this->params['access_key_id'],$this->params['access_key_secret'],$this->params['bucket'],$this->params['url']);

        $filename = $this->params['filename'] ?: $this->random;
        if (strrpos($filename, '.')) {
            $filename = substr($filename, 0, strrpos($filename, '.'));
        }
        if ($this->params['terminal']) {
            if(!preg_match("/^[a-z]+$/", $this->params['terminal'])) {
                throw new \Exception('terminal format error');
            }
            $this->params['terminal'] = $this->params['terminal'] . '/';
        }
        $filename     = 'download/' . $this->params['terminal'] . trim($filename, '/.');
        $localFileExt = pathinfo($localFile, PATHINFO_EXTENSION);
        $fileOssUrl   = $s->uploadFile($filename . '.' . $localFileExt, $localFile);

        unlink($localFile);

        if($this->ext == '.csv' && is_dir($this->params['upload_temp_dir'] . $this->random)) {
            rmdir($this->params['upload_temp_dir'] . $this->random);
        }

        return $fileOssUrl;
    }
}