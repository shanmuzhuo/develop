<?php

namespace app\index\controller;

use think\Controller;
use think\Loader;

/**
 * 发送电子邮件的demo：
 */

Loader::import('PhpOffice/PHPExcel/Classes/PHPExcel', EXTEND_PATH);
Loader::import('PhpOffice/PHPExcel/Classes/PHPExcel/IOFactory.php', EXTEND_PATH);

class Handle extends Controller
{
    public function dohandle()
    {
        $comparr = $this->handleComp(); //获取公司考勤数据
        $factarr = $this->handFactory(); //获取工厂考勤数据

        // 定义公司，工厂的结果
        $compresult = array();
        $factresult = array();

        /*
         * 处理两个二维数组，
         * 目标： 1.获取在公司，工厂都有考勤数据的人员信息，替换人员的信息
         * 方式： 1，直接在数组里面处理，2，将特殊情况独拎出来在集中处理，选择2
         * */

        // 要留在公司的人员名单
        $staycomp = array("王楚斌", "苏敏", "陈俊安", "蒙其伟", "张升", "胡席林", "袁钰", "李世昌", "阮昱臻", "褚应平", "莫国锋");
        $comptemp = array();
        $facttemp = array();

        // 将莫国峰，褚应平的信息单独拎出来
        $ussetarry = array("褚应平", "莫国锋");
        $ussetarrytemp = array();

        // 获取工厂特殊人员的信息：
        $count2 = 0;
        foreach ($factarr as $item) {
            foreach ($staycomp as $scp) {
                if ($item['name'] == $scp) {
                    array_push($facttemp, $item);
                    unset($factarr[$count2]);
                }
            }
            $count2++;
        }

        // 获取公司特殊人员的信息：
        $count = 0;
        foreach ($comparr as $item) {
            foreach ($staycomp as $scp) {
                if ($item['name'] == $scp) {
                    array_push($comptemp, $item);
                    unset($comparr[$count]);
                }
            }
            $count++;
        }

        // 处理留在公司or工厂的人员信息
        foreach ($comptemp as &$item1) {
            foreach ($facttemp as &$item2) {
                $date1 = explode(" ", $item1['date'])[0];
                $date2 = str_replace("/", "-", $item2['date']); //比较日期
                if ($date1 == $date2 && $item1['name'] == $item2['name']) {
                    // 比较签到时间
                    if ($item1['starttime'] != NULL && $item2['starttime'] != NUll) {
                        $stime1 = explode(":", $item1['starttime']);
                        $stime1 = mktime(intval($stime1[0]), intval($stime1[1]));

                        $stime2 = explode(":", $item2['starttime']);
                        $stime2 = mktime(intval($stime2[0]), intval($stime2[1])); // 设置时间

                        if ($stime1 > $stime2) {
                            $item1['starttime'] = $item2['starttime'];
                        }
                    } else if ($item1['starttime'] == NULL && $item2['starttime'] != NUll) {
                        $item1['starttime'] = $item2['starttime'];
                    }

                    // 比较签退步时间
                    if ($item1['endtime'] != NULL && $item2['endtime'] != NUll) {
                        $stime1 = explode(":", $item1['endtime']);
                        $stime1 = mktime(intval($stime1[0]), intval($stime1[1]));

                        $stime2 = explode(":", $item2['endtime']);
                        $stime2 = mktime(intval($stime2[0]), intval($stime2[1])); // 设置时间

                        if ($stime1 < $stime2) {
                            $item1['endtime'] = $item2['endtime'];
                        }
                    } else if ($item1['endtime'] == NULL && $item2['endtime'] != NUll) {
                        $item1['endtime'] = $item2['endtime'];
                    }
                    break;
                }
            }
        }

        $count3 = 0;
        foreach ($comptemp as $item) {
            foreach ($ussetarry as $item2) {
                if ($item['name'] == $item2) {
                    array_push($ussetarrytemp, $item);
                    unset($comptemp[$count3]);
                }
            }
            $count3++;
        }

        $compresult = array_merge($comparr, $comptemp);
        $factresult = array_merge($factarr, $ussetarrytemp);

        /*
         * 生成公司的考勤模板最终章
         * */

        error_reporting(E_ALL);
        ini_set('display_errors', TRUE);
        ini_set('display_startup_errors', TRUE);
        date_default_timezone_set('Europe/London');
        $objPHPExcel = new \PHPExcel();

        // 设置文档属性：
        $objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
            ->setLastModifiedBy("Maarten Balliauw")
            ->setTitle("PHPExcel Test Document")
            ->setSubject("PHPExcel Test Document")
            ->setDescription("Test document for PHPExcel, generated using PHP classes.")
            ->setKeywords("office PHPExcel php")
            ->setCategory("Test result file");

        // 设置第一,二行数据
        $objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', '登记号码')
            ->setCellValue('B1', '姓名')
            ->setCellValue('C1', '日期')
            ->setCellValue('D1', '班次')
            ->setCellValue('E1', '上班时间')
            ->setCellValue('F1', '下班时间')
            ->setCellValue('G1', '签到时间')
            ->setCellValue('H1', '签退时间')
            ->setCellValue('I1', '迟到分钟')
            ->setCellValue('J1', '早退分钟')
            ->setCellValue('I2', 'FALSE')
            ->setCellValue('J2', 'FALSE')
            ->setCellValue('K1', '部门班组');

        // 设置表身的数据
        for ($i = 0; $i < count($compresult); $i++) {
            $ul = 3 + $i;
            $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('A' . $ul, $compresult[$i]['num'])
                ->setCellValue('B' . $ul, $compresult[$i]['name'])
                ->setCellValue('C' . $ul, $compresult[$i]['date'])
                ->setCellValue('D' . $ul, intval($compresult[$i]['banci']) < 4 ? $compresult[$i]['banci'] : "")
                ->setCellValue('G' . $ul, $compresult[$i]['starttime'])
                ->setCellValue('H' . $ul, $compresult[$i]['endtime'])
                ->setCellValue('I' . $ul, 0)
                ->setCellValue('J' . $ul, 0)
                ->setCellValue('K' . $ul, $compresult[$i]['dept']);

            // 班次对应的时间迟到分钟
            if (intval($compresult[$i]['banci']) == 1) {
                $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('E' . $ul, '9:30')
                    ->setCellValue('F' . $ul, '19:00');

                $this->deal_time('9:30', '19:00', $compresult[$i]['starttime'], $compresult[$i]['endtime'], $objPHPExcel, $ul);
            }
            if (intval($compresult[$i]['banci']) == 2) {
                $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('E' . $ul, '9:00')
                    ->setCellValue('F' . $ul, '18:10');

                $this->deal_time('9:00', '18:10', $compresult[$i]['starttime'], $compresult[$i]['endtime'], $objPHPExcel, $ul);

            }
            if (intval($compresult[$i]['banci']) == 3) {
                $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('E' . $ul, '10:00')
                    ->setCellValue('F' . $ul, '19:30');

                $this->deal_time('10:00', '19:30', $compresult[$i]['starttime'], $compresult[$i]['endtime'], $objPHPExcel, $ul);
            }

            if (intval($compresult[$i]['banci']) == 4) {
                $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('E' . $ul, '9:00')
                    ->setCellValue('F' . $ul, '18:30');

                $this->deal_time('9:00', '18:30', $compresult[$i]['starttime'], $compresult[$i]['endtime'], $objPHPExcel, $ul);
            }

        }
        /*        // Rename worksheet
                $objPHPExcel->getActiveSheet()->setTitle('Simple');


        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
                $objPHPExcel->setActiveSheetIndex(0);

                ob_end_clean();
                ob_start();
        // Redirect output to a client’s web browser (Excel2007)
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                header('Content-Disposition: attachment;filename="01simple.xlsx"');
                header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
                header('Cache-Control: max-age=1');

        // If you're serving to IE over SSL, then the following may be needed
                header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
                header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
                header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
                header('Pragma: public'); // HTTP/1.0

                $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
                $objWriter->save('php://output');*/


        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save(ROOT_PATH . "public" . DS . "uploads/excle/公司考勤结果.xlsx"); //生成公司的考勤结果


        /*
         * 生成工厂的考勤模板最终章
         * */
        $objPHPExcel2 = new \PHPExcel();

        // 设置文档属性：
        $objPHPExcel2->getProperties()->setCreator("Maarten Balliauw")
            ->setLastModifiedBy("Maarten Balliauw")
            ->setTitle("PHPExcel Test Document")
            ->setSubject("PHPExcel Test Document")
            ->setDescription("Test document for PHPExcel, generated using PHP classes.")
            ->setKeywords("office PHPExcel php")
            ->setCategory("Test result file");

        // 设置第一,二行数据
        $objPHPExcel2->setActiveSheetIndex(0)
            ->setCellValue('A1', '登记号码')
            ->setCellValue('B1', '姓名')
            ->setCellValue('C1', '日期')
            ->setCellValue('D1', '班次')
            ->setCellValue('E1', '上班时间')
            ->setCellValue('F1', '下班时间')
            ->setCellValue('G1', '签到时间')
            ->setCellValue('H1', '签退时间')
            ->setCellValue('I1', '迟到分钟')
            ->setCellValue('J1', '早退分钟')
            ->setCellValue('I2', 'FALSE')
            ->setCellValue('J2', 'FALSE')
            ->setCellValue('K1', '部门班组');

        // 设置表身的数据
        for ($i = 0; $i < count($factresult); $i++) {
            $ul = 3 + $i;
            $objPHPExcel2->setActiveSheetIndex(0)
                ->setCellValue('A' . $ul, $factresult[$i]['num'])
                ->setCellValue('B' . $ul, $factresult[$i]['name'])
                ->setCellValue('C' . $ul, $factresult[$i]['date'])
                ->setCellValue('D' . $ul, "")
                ->setCellValue('E' . $ul, "8:30")
                ->setCellValue('F' . $ul, "18:00")
                ->setCellValue('G' . $ul, $factresult[$i]['starttime'])
                ->setCellValue('H' . $ul, $factresult[$i]['endtime'])
                ->setCellValue('I' . $ul, 0)
                ->setCellValue('J' . $ul, 0)
                ->setCellValue('K' . $ul, $factresult[$i]['dept']);


            // 设置上班时间，下班时间
            // 1、特殊的人员有：彭贝：财务中心（9.30-19.30）
            // 2、杨旗：夜班（20.30-6.30）
            // 3、一般职员：9.00-18.30
            if ($factresult[$i]['name'] == "彭贝") {
                $objPHPExcel2->setActiveSheetIndex(0)
                    ->setCellValue('E' . $ul, "9:30")
                    ->setCellValue('F' . $ul, "19:00");

                $this->deal_time("9:30", "19:00", $factresult[$i]['starttime'], $factresult[$i]['endtime'], $objPHPExcel2, $ul);
            } else if ($factresult[$i]['name'] == "杨旗") {
                $objPHPExcel2->setActiveSheetIndex(0)
                    ->setCellValue('E' . $ul, "20:30")
                    ->setCellValue('F' . $ul, "6:30");

                $this->deal_time("20:30", "6:30", $factresult[$i]['starttime'], $factresult[$i]['endtime'], $objPHPExcel2, $ul);
            } else {
                $this->deal_time('8:30', '18:00', $factresult[$i]['starttime'], $factresult[$i]['endtime'], $objPHPExcel2, $ul);
            }

            // 将形如2017-11-27 一 改为 2017/11/27
            if ($factresult[$i]['name'] == "莫国锋" || $factresult[$i]['name'] == "褚应平") {
                $temp = substr($factresult[$i]['date'], 0, 10);
                $temp = explode("-", $temp);
                $strdate = $temp[0] . "/" . "$temp[1]" . "/" . $temp[2];
                $objPHPExcel2->setActiveSheetIndex(0)
                    ->setCellValue('C' . $ul, $strdate);
            }

        }
        /*        ob_end_clean();
                ob_start();
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                header('Content-Disposition: attachment;filename="工厂考勤处理结果.xlsx"');
                header('Cache-Control: max-age=0');
                header('Cache-Control: max-age=1');

                header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
                header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
                header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
                header('Pragma: public'); // HTTP/1.0

                $objWriter2 = \PHPExcel_IOFactory::createWriter($objPHPExcel2, 'Excel2007');
                $objWriter2->save('php://output');*/

        $objWriter2 = \PHPExcel_IOFactory::createWriter($objPHPExcel2, 'Excel2007');
        $objWriter2->save(ROOT_PATH . "public" . DS . "uploads/excle/工厂考勤结果.xlsx"); //生成工厂的考勤结果

        $this->addFileToZip();

    }

    // 处理公司的考勤表
    public function handleComp()
    {
        // Check prerequisites
        if (!file_exists(ROOT_PATH . 'public' . DS . 'uploads/excle/complist.xls')) {
            exit("未找到文件：" . ROOT_PATH . "public" . DS . "uploads/excle/complist.xls\n");
        }

        $reader = \PHPExcel_IOFactory::createReader('Excel5'); //设置以Excel5格式(Excel97-2003工作簿)
        $PHPExcel = $reader->load(ROOT_PATH . 'public' . DS . 'uploads/excle/complist.xls'); // 载入excel文件
        $sheet = $PHPExcel->getSheet(0); // 读取第一個工作表
        $highestRow = $sheet->getHighestRow(); // 取得总行数
        $highestColumm = $sheet->getHighestColumn(); // 取得总列数

        $columns = array("部门班组", "用户编号", "姓名", "日期", "上班时间1", "下班时间1", "班次");
        $deleteColums = array(); // 要删除数组

        for ($column = 'A'; $column <= $highestColumm; $column++) {//列数是以A列开始
            if (!in_array($sheet->getCell($column . "1")->getValue(), $columns)) {
                array_push($deleteColums, $column); // 将要删除的列添加进入一个删除数组
            }
        }

        // 要按照倒顺删除，不然会删除错误的列。
        for ($i = count($deleteColums) - 1; $i >= 0; $i--) {
            $sheet->removeColumn("" . $deleteColums[$i], 1);
        }

        /** 循环读取每个单元格的数据,按照班次顺序排序 */
        $arrry = array();
        $toalarrry = array();
        for ($row = 2; $row <= $highestRow; $row++) {
            for ($column = 'A'; $column <= $highestColumm; $column++) {//列数是以A列开始
                $columnName = $sheet->getCell($column . "1")->getValue();
                $columnValue = $sheet->getCell($column . $row)->getValue();
                if ($columnName == "姓名") {
                    $arrry["name"] = $columnValue;
                }
                if ($columnName == "班次") {
                    if ($columnValue == NUll) {
                        $arrry["banci"] = 4.0;
                    } else {
                        $arrry["banci"] = $columnValue;
                    }
                }
                if ($columnName == "部门班组") {
                    $arrry["dept"] = $columnValue;
                }
                if ($columnName == "日期") {
                    $arrry["date"] = $columnValue;
                }
                if ($columnName == "上班时间1") {
                    $arrry["starttime"] = $columnValue;
                }
                if ($columnName == "下班时间1") {
                    $arrry["endtime"] = $columnValue;
                }
                if ($columnName == "用户编号") {
                    $arrry["num"] = $columnValue;
                }
            }
            array_push($toalarrry, $arrry);
        }


        $ages = array();
        foreach ($toalarrry as $user) {
            $banci[] = $user['banci'];
        }
        array_multisort($banci, SORT_ASC, $toalarrry); // 按照班次进行排序
        return $toalarrry;

    }

    // 处理工厂的考勤表
    public function handFactory()
    {
        // Check prerequisites
        if (!file_exists(ROOT_PATH . 'public' . DS . 'uploads/excle/factorylist2.xls')) {
            exit("未找到文件：" . ROOT_PATH . "public" . DS . "uploads/excle/factorylist2.xls\n");
        }

        $reader = \PHPExcel_IOFactory::createReader('Excel5'); //设置以Excel5格式(Excel97-2003工作簿)
        $PHPExcel = $reader->load(ROOT_PATH . 'public' . DS . 'uploads/excle/factorylist2.xls'); // 载入excel文件

        $sheet = $PHPExcel->getSheet(0); // 读取第一個工作表
        $highestRow = $sheet->getHighestRow(); // 取得总行数
        $highestColumm = $sheet->getHighestColumn(); // 取得总列数
        ++$highestColumm;

        $columns = array("自定义编号", "姓名", "日期", "班次", "上班时间", "下班时间", "签到时间", "签退时间", "部门");
        $deleteColums = array(); // 要删除数组

        for ($column = 'A'; $column != $highestColumm; $column++) {//列数是以A列开始
            if (!in_array($sheet->getCell($column . "1")->getValue(), $columns)) {
                array_push($deleteColums, $column); // 将要删除的列添加进入一个删除数组
            }
        }

        // 要按照倒顺删除，不然会删除错误的列。
        for ($i = count($deleteColums) - 1; $i >= 0; $i--) {
            $znum = \PHPExcel_Cell::columnIndexFromString("Z");
            if ($deleteColums[$i] < "Z") {
                $sheet->removeColumn("" . $deleteColums[$i], 1);
            }
        }

        /** 循环读取每个单元格的数据,按照班次顺序排序 */
        $arrry = array();
        $toalarrry = array();
        for ($row = 2; $row <= $highestRow; $row++) {
            for ($column = 'A'; $column != $highestColumm; $column++) {//列数是以A列开始
                $columnName = $sheet->getCell($column . "1")->getValue();
                $columnValue = $sheet->getCell($column . $row)->getValue();
                if ($columnName == "姓名") {
                    $arrry["name"] = $columnValue;
                }
                if ($columnName == "部门") {
                    $arrry["dept"] = $columnValue;
                }
                if ($columnName == "日期") {
                    $arrry["date"] = $columnValue;
                }
                if ($columnName == "签到时间") {
                    $arrry["starttime"] = $columnValue;
                }
                if ($columnName == "签退时间") {
                    $arrry["endtime"] = $columnValue;
                }
                if ($columnName == "自定义编号") {
                    $arrry["num"] = $columnValue;
                }
            }
            array_push($toalarrry, $arrry);
        }

        $ages = array();
        foreach ($toalarrry as $user) {
            $num[] = $user['num'];
        }
        array_multisort($num, SORT_ASC, $toalarrry); // 按照自定义编号进行排序

        $newarray = array();
        // 按照自定义编号删除无用的数据
        for ($m = 0; $m < count($toalarrry); $m++) {
            $delnum = strval($toalarrry[$m]["num"]);
            if (($delnum == "40") || (strlen($delnum) == "4" && ord($delnum) != "57" || $delnum == "10033" || $delnum == "10107")) {
//                $newarray =  array_slice($toalarrry,$m,1);
                $newarray = array_merge_recursive($newarray, array_slice($toalarrry, $m, 1));
            }
        }

        return $newarray;
    }

    /**
     * 用来处理上下班时间，迟到早退的时间
     * @param $begaintiem 规定上班时间
     * @param $stoptime 规定下班时间
     * @param $starttime 真正开始上班时间
     * @param $endtime 真正开始下班时间
     * @param $objPHPExcel 要操作的Excle对象
     * @param $uls 要填写行数列
     */
    function deal_time($begaintiem, $stoptime, $starttime, $endtime, $objPHPExcel, $uls)
    {
        //迟到
        if ($starttime != null) {
            $stime1 = explode(":", $begaintiem);
            $stime1 = mktime(intval($stime1[0]), intval($stime1[1]));

            $stime2 = explode(":", $starttime);
            $stime2 = mktime(intval($stime2[0]), intval($stime2[1])); // 设置时间
            $difftime1 = $stime2 - $stime1; // 获取迟到时间差,大于0为迟到

            if ($difftime1 > 0) {
                $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('I' . $uls, $difftime1 / 60);
            }

        } else {
            $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('I' . $uls, 'FALSE');
        }

        //早退
        if ($endtime != null) {
            $stime3 = explode(":", $stoptime);
            $stime3 = mktime(intval($stime3[0]), intval($stime3[1]));

            $stime4 = explode(":", $endtime);
            $stime4 = mktime(intval($stime4[0]), intval($stime4[1])); // 设置时间
            $difftime2 = $stime3 - $stime4; // 获取早退时间差,大于0为早退

            if ($difftime2 > 0 && $difftime2 < 60000) {
                $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('J' . $uls, $difftime2 / 60);
            }
        } else {
            $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('J' . $uls, 'FALSE');
        }
    }

    // 生成压缩包，并下载
    function addFileToZip()
    {
        $fileList = array(
            ROOT_PATH . "public" . DS . "uploads/excle/公司考勤结果.xlsx",
            ROOT_PATH . "public" . DS . "uploads/excle/工厂考勤结果.xlsx"
        );
        $filename = "test.zip";
        $zip = new \ZipArchive();
        $zip->open($filename, \ZipArchive::CREATE);   //打开压缩包
        foreach ($fileList as $file) {
            if (!file_exists($file)) {
                exit("未找到文件：" .$file);
            }
            $zip->addFile($file, basename($file));   //向压缩包中添加文件
        }
        echo "压缩文件";
        $zip->close();  //关闭压缩包
        ob_end_clean();
        ob_start();
        header("Cache-Control: public");
        header("Content-Description: File Transfer");
        header('Content-disposition: attachment; filename='.basename($filename)); //文件名
        header("Content-Type: application/zip"); //zip格式的
        header("Content-Transfer-Encoding: binary"); //告诉浏览器，这是二进制文件
        header('Content-Length: '. filesize($filename)); //告诉浏览器，文件大小
        @readfile($filename);
    }

}