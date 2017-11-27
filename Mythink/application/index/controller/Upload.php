<?php

namespace app\index\controller;

use think\Controller;
use think\File;

class Upload extends Controller
{
    public function upload()
    {
        // 获取表单上传文件
        $files = request()->file('xlsx');
        $flag = 0;
        $count = 1;
        foreach ($files as $file) {
            // 移动到框架应用根目录/public/uploads/ 目录下
            if ($file->checkExt(array("xlsx", "xls"))) {
                if ($count == 1) {
                    $info = $file->move(ROOT_PATH . 'public' . DS . 'uploads/excle', "complist"); //公司
                }
                if ($count == 2) {
                    $info = $file->move(ROOT_PATH . 'public' . DS . 'uploads/excle', "factorylist2"); //工厂
                }

                if ($info) {
                    // 成功上传后 获取上传信息
                    $flag ++ ;
                } else {
                    // 上传失败获取错误信息
                    echo $file->getError();
                }
            } else {
                echo "请上传正确的文件格式";
            }

            $count++;
        }

        if($flag == 2){
            echo "$count";
            echo "上传成功开始处理文件";
            $handle = new Handle();
            $handle->dohandle();
        }
    }

    /**
     * 跳转到上传文件的界面，多文件上传
     */
    public function doupload()
    {
        return $this->fetch();
    }

}

