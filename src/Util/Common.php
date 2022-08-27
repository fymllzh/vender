<?php
/*
 * @Description: 公共方法类
 * @version: 1.0
 * @Author: zhihao.li
 * @Date: 2022-1-24 17:44:17
 */

namespace FFMBase\Util;

use WebGeeker\Validation\Validation;

class Common
{
    /**
     * @desc 参数验证
     * @param $data
     * @param $validator
     */
    public static function validator($data, $validator)
    {
        $errors = [];
        foreach ($validator as $k => $item) {
            try {
                Validation::validate($data, [$k => $item]);
            } catch (\Exception $e) {
                $errors[$k] = $e->getMessage();
            }
        }

        if ($errors) {
            throw new \Exception( json_encode($errors, JSON_UNESCAPED_UNICODE));
        }
    }
}