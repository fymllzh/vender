<?php
/*
 * @Description: 导出入口
 * @version: 1.0
 * @Author: panjia
 * @Date: 2022-1-13 20:44:17
 */

namespace FFMBase\Export;

class Export
{
    private $export = NULL;

    const ARR = [
        'type'              => 'excel', //导出类型
        'upload_temp_dir'   => '/tmp/', //临时目录
        'filename'          => '', //文件名字
        'terminal'          => '', //终端来源标识
        'oss'               => false, //是否上传阿里云
        'url'               => '', //阿里云文件访问地址
        'endpoint'          => '',
        'access_key_id'     => '',
        'access_key_secret' => '',
        'bucket'            => '',
    ];

    public function __construct(array $arr = [])
    {
        $params = $arr + self::ARR;

        $type = strtolower($params['type']);
        if (!in_array($type, ['excel', 'csv'])) {
            throw new \Exception('error type');
        }

        $params['upload_temp_dir'] = rtrim($params['upload_temp_dir'],'/') . '/';
        if(!is_dir($params['upload_temp_dir']) && !mkdir($params['upload_temp_dir'], 0777,true)) {
            throw new \Exception('folder does not exist or failed to create');
        }

        $type = ucfirst($type);
        $class = '\\FFMBase\\Export\\'.$type;
        $this->export = new $class($params);
    }

    public function __call($methodName, $arguments)
    {
        if(!method_exists($this->export,$methodName)) {
            throw new \Exception('method not exist');
        }

        return call_user_func_array([$this->export,$methodName],$arguments);
    }

}