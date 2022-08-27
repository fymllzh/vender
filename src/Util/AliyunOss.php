<?php
/*
 * @Description: Oss基础类
 * @version: 1.0
 * @Author: zhihao.li
 * @Date: 2022-1-21 14:44:17
 */

namespace FFMBase\Util;

use OSS\OssClient;

class AliyunOss
{
    protected $ossClient;
    protected $bucket;
    protected $url;

    public function __construct($endpoint,$accessKeyId,$accessKeySecret,$bucket,$url)
    {
        try {
            $ossClient = new OssClient($accessKeyId, $accessKeySecret, $endpoint, false);
        } catch (\Exception $e) {
            throw new \Exception('oss bucket not exist');
        }
        $this->ossClient = $ossClient;
        $this->bucket = $bucket;
        $this->url = $url;
    }

    /**
     * 上传文件
     * @param null $object 文本路径机文件明 例如：base64/aa.txt 不能以/开头
     * @param null $localFileUrl 本地文件路径
     * @param bool $isCheckExists 检验是否已添加
     * @return Array
     * @throws Exception
     */
    public function uploadFile($object = null, $localFileUrl = null, $isCheckExists = false)
    {
        if (!$localFileUrl || !$object) {
            throw new \Exception('parameter error');
        }

        if ($isCheckExists && $this->ossClient->doesObjectExist($this->bucket, $object)) {
            throw new \Exception('file has been added');
        }

        $this->ossClient->uploadFile($this->bucket, $object, $localFileUrl);

        if ($this->ossClient->doesObjectExist($this->bucket, $object)) {
            return $this->url . $object;
        }

        throw new \Exception('file save failed');
    }
}