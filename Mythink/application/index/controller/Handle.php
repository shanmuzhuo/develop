<?php
namespace app\index\controller;
use think\Config;
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
         * 方式： 1，直接在数组里面处理，2，将特殊情况独拎出来在集中处理，选择2，理由不详。。。
         * */

        // 要留在公司的人员名单
        $staycomp = array("王楚斌","苏敏","陈俊安","蒙其伟","张升","胡席林","袁钰","李世昌","阮昱臻","褚应平","莫国锋");
        $comptemp = array();
        $facttemp = array();

        // 获取工厂特殊人员的信息：
        foreach ($factarr as $item) {
            foreach ($staycomp as $scp) {
                if($item['name'] == $scp){
                    array_push($facttemp, $item);
                }
            }
        }

        // 获取公司特殊人员的信息：
        $count = 0;
        foreach ($comparr as $item) {
            foreach ($staycomp as $scp) {
                if($item['name'] == $scp){
                    array_push($comptemp, $item);
                    unset($comparr[$count]);
                }
            }
            $count++;
        }

        // 处理留在公司的人员
        foreach ($comptemp as &$item1) {
            foreach ($facttemp as &$item2) {
                $date1 = explode(" ",$item1['date'])[0] ;
                $date2 = str_replace("/","-",$item2['date']); //比较日期
                if($date1 == $date2 && $item1['name'] == $item2['name']){
                    // 比较签到时间
                    if($item1['starttime'] != NULL && $item2['starttime'] != NUll){
                        $stime1 = explode(":", $item1['starttime']);
                        $stime1 = mktime(intval($stime1[0]),intval($stime1[1]));

                        $stime2 = explode(":", $item2['starttime']);
                        $stime2 = mktime(intval($stime2[0]),intval($stime2[1])); // 设置时间

                        if($stime1 > $stime2){
                            $item1['starttime'] = $item2['starttime'];
                        }
                    }else if($item1['starttime'] == NULL && $item2['starttime'] != NUll){
                        $item1['starttime'] = $item2['starttime'];
                    }

                    // 比较签退步时间
                    if($item1['endtime'] != NULL && $item2['endtime'] != NUll){
                        $stime1 = explode(":", $item1['endtime']);
                        $stime1 = mktime(intval($stime1[0]),intval($stime1[1]));

                        $stime2 = explode(":", $item2['endtime']);
                        $stime2 = mktime(intval($stime2[0]),intval($stime2[1])); // 设置时间

                        if($stime1 < $stime2){
                            $item1['endtime'] = $item2['endtime'];
                        }
                    }else if($item1['endtime'] == NULL && $item2['endtime'] != NUll){
                        $item1['endtime'] = $item2['endtime'];
                    }
                    break;
                }
            }
        }

        $compresult  = array_merge($comparr,$comptemp);
//        echo dump($comptemp);
//        echo dump($comparr);
//        echo dump($compresult);
//        echo dump($facttemp);

        /*
         * 生成公司的考勤模板最终章
         * */
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
        for ($i = 0 ;$i <count($compresult);$i++){
            $ul = 3 + $i;
            $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('A'.$ul, $compresult[$i]['num'])
                ->setCellValue('B'.$ul, $compresult[$i]['name'])
                ->setCellValue('C'.$ul, $compresult[$i]['date'])
                ->setCellValue('D'.$ul, intval($compresult[$i]['banci'])<4?$compresult[$i]['banci']:"")
                ->setCellValue('G'.$ul, $compresult[$i]['starttime'])
                ->setCellValue('H'.$ul, $compresult[$i]['endtime'])
                ->setCellValue('I'.$ul, 0)
                ->setCellValue('J'.$ul, 0)
                ->setCellValue('K'.$ul, $compresult[$i]['dept']);

            // 班次对应的时间迟到分钟
            if(intval($compresult[$i]['banci']) == 1){
                $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('E'.$ul, '9:30')
                ->setCellValue('F'.$ul, '19:00');

                //迟到
                if($compresult[$i]['starttime'] != null){
                    $stime1 = explode(":", '9:30');
                    $stime1 = mktime(intval($stime1[0]),intval($stime1[1]));

                    $stime2 = explode(":", $compresult[$i]['starttime']);
                    $stime2 = mktime(intval($stime2[0]),intval($stime2[1])); // 设置时间
                    $difftime1 = $stime2 - $stime1; // 获取迟到时间差,大于0为迟到

                    if($difftime1 > 0 ){
                        $objPHPExcel->setActiveSheetIndex(0)
                            ->setCellValue('I'.$ul, $difftime1/60);
                    }

                }else{
                    $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('I'.$ul, 'FALSE');
                }

                //早退
                if($compresult[$i]['endtime'] != null){
                    $stime3 = explode(":", '19:00');
                    $stime3 = mktime(intval($stime3[0]),intval($stime3[1]));

                    $stime4 = explode(":", $compresult[$i]['endtime']);
                    $stime4 = mktime(intval($stime4[0]),intval($stime4[1])); // 设置时间
                    $difftime2 = $stime3 - $stime4; // 获取早退时间差,大于0为早退

                    if($difftime2> 0 ){
                        $objPHPExcel->setActiveSheetIndex(0)
                            ->setCellValue('J'.$ul, $difftime2/60);
                    }
                }else{
                    $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('J'.$ul, 'FALSE');
                }


            }
            if(intval($compresult[$i]['banci']) == 2){
                $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('E'.$ul, '9:00')
                    ->setCellValue('F'.$ul, '18:10');

                if($compresult[$i]['starttime'] != null){
                    $stime1 = explode(":", '9:00');
                    $stime1 = mktime(intval($stime1[0]),intval($stime1[1]));

                    $stime2 = explode(":", $compresult[$i]['starttime']);
                    $stime2 = mktime(intval($stime2[0]),intval($stime2[1])); // 设置时间
                    $difftime1 = $stime2 - $stime1; // 获取迟到时间差,大于0为迟到

                    if($difftime1 > 0 ){
                        $objPHPExcel->setActiveSheetIndex(0)
                            ->setCellValue('I'.$ul, $difftime1/60);
                    }

                }else{
                    $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('I'.$ul, 'FALSE');
                }

                if($compresult[$i]['endtime'] != null){
                    $stime3 = explode(":", '18:10');
                    $stime3 = mktime(intval($stime3[0]),intval($stime3[1]));

                    $stime4 = explode(":", $compresult[$i]['endtime']);
                    $stime4 = mktime(intval($stime4[0]),intval($stime4[1])); // 设置时间
                    $difftime2 = $stime3 - $stime4; // 获取早退时间差,大于0为早退

                    if($difftime2> 0 ){
                        $objPHPExcel->setActiveSheetIndex(0)
                            ->setCellValue('J'.$ul, $difftime2/60);
                    }
                }else{
                    $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('J'.$ul, 'FALSE');
                }
            }
            if(intval($compresult[$i]['banci']) == 3){
                $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('E'.$ul, '10:00')
                    ->setCellValue('F'.$ul, '19:30');

                if($compresult[$i]['starttime'] != null){
                    $stime1 = explode(":", '10:00');
                    $stime1 = mktime(intval($stime1[0]),intval($stime1[1]));

                    $stime2 = explode(":", $compresult[$i]['starttime']);
                    $stime2 = mktime(intval($stime2[0]),intval($stime2[1])); // 设置时间
                    $difftime1 = $stime2 - $stime1; // 获取迟到时间差,大于0为迟到

                    if($difftime1 > 0 ){
                        $objPHPExcel->setActiveSheetIndex(0)
                            ->setCellValue('I'.$ul, $difftime1/60);
                    }

                }else{
                    $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('I'.$ul, 'FALSE');
                }

                if($compresult[$i]['endtime'] != null){
                    $stime3 = explode(":", '19:30');
                    $stime3 = mktime(intval($stime3[0]),intval($stime3[1]));

                    $stime4 = explode(":", $compresult[$i]['endtime']);
                    $stime4 = mktime(intval($stime4[0]),intval($stime4[1])); // 设置时间
                    $difftime2 = $stime3 - $stime4; // 获取早退时间差,大于0为早退

                    if($difftime2> 0 ){
                        $objPHPExcel->setActiveSheetIndex(0)
                            ->setCellValue('J'.$ul, $difftime2/60);
                    }
                }else{
                    $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('J'.$ul, 'FALSE');
                }

            }

            if(intval($compresult[$i]['banci']) == 4){
                $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('E'.$ul, '9:00')
                    ->setCellValue('F'.$ul, '18:30');

                if($compresult[$i]['starttime'] != null){
                    $stime1 = explode(":", '9:00');
                    $stime1 = mktime(intval($stime1[0]),intval($stime1[1]));

                    $stime2 = explode(":", $compresult[$i]['starttime']);
                    $stime2 = mktime(intval($stime2[0]),intval($stime2[1])); // 设置时间
                    $difftime1 = $stime2 - $stime1; // 获取迟到时间差,大于0为迟到

                    if($difftime1 > 0 ){
                        $objPHPExcel->setActiveSheetIndex(0)
                            ->setCellValue('I'.$ul, $difftime1/60);
                    }

                }else{
                    $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('I'.$ul, 'FALSE');
                }

                if($compresult[$i]['endtime'] != null){
                    $stime3 = explode(":", '18:30');
                    $stime3 = mktime(intval($stime3[0]),intval($stime3[1]));

                    $stime4 = explode(":", $compresult[$i]['endtime']);
                    $stime4 = mktime(intval($stime4[0]),intval($stime4[1])); // 设置时间
                    $difftime2 = $stime3 - $stime4; // 获取早退时间差,大于0为早退

                    if($difftime2> 0 ){
                        $objPHPExcel->setActiveSheetIndex(0)
                            ->setCellValue('J'.$ul, $difftime2/60);
                    }
                }else{
                    $objPHPExcel->setActiveSheetIndex(0)
                        ->setCellValue('J'.$ul, 'FALSE');
                }
            }

        }

        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save(ROOT_PATH . "public" . DS . "uploads/excle/compresult.xlsx");

	}

	// 处理公司的考勤表 
	public function handleComp()
	{
		// Check prerequisites
		if (!file_exists(ROOT_PATH . 'public' . DS . 'uploads/excle/complist.xls')) {
			exit("未找到文件：".ROOT_PATH . "public" . DS . "uploads/excle/complist.xls\n");
		}

		$reader = \PHPExcel_IOFactory::createReader('Excel5'); //设置以Excel5格式(Excel97-2003工作簿)
		$PHPExcel = $reader->load(ROOT_PATH . 'public' . DS . 'uploads/excle/complist.xls'); // 载入excel文件
		$sheet = $PHPExcel->getSheet(0); // 读取第一個工作表
		$highestRow = $sheet->getHighestRow(); // 取得总行数
		$highestColumm = $sheet->getHighestColumn(); // 取得总列数

        $columns = array("部门班组","用户编号","姓名","日期","上班时间1","下班时间1","班次");
        $deleteColums = array(); // 要删除数组

        for ($column = 'A'; $column <= $highestColumm; $column++) {//列数是以A列开始
            if(!in_array($sheet->getCell($column."1")->getValue(),$columns)){
                array_push($deleteColums,$column); // 将要删除的列添加进入一个删除数组
            }
        }

        // 要按照倒顺删除，不然会删除错误的列。
        for ($i = count($deleteColums)-1 ;$i >= 0;$i--){
            $sheet->removeColumn("".$deleteColums[$i],1);
        }

        /** 循环读取每个单元格的数据,按照班次顺序排序 */
        $arrry = array();
        $toalarrry = array();
        for ($row = 2 ; $row <= $highestRow ; $row ++){
            for ($column = 'A'; $column <= $highestColumm; $column++) {//列数是以A列开始
                $columnName = $sheet->getCell($column."1")->getValue();
                $columnValue = $sheet->getCell($column.$row)->getValue();
                if( $columnName == "姓名"){
                    $arrry["name"] = $columnValue;
                }
                if( $columnName == "班次"){
                    if($columnValue == NUll){
                        $arrry["banci"] = 4.0;
                    }else{
                        $arrry["banci"] = $columnValue;
                    }
                }
                if( $columnName == "部门班组"){
                    $arrry["dept"] = $columnValue;
                }
                if( $columnName == "日期"){
                    $arrry["date"] = $columnValue;
                }
                if( $columnName == "上班时间1"){
                    $arrry["starttime"] = $columnValue;
                }
                if( $columnName == "下班时间1"){
                    $arrry["endtime"] = $columnValue;
                }
                if( $columnName == "用户编号"){
                    $arrry["num"] = $columnValue;
                }
            }
            array_push($toalarrry ,$arrry);
        }


//        echo "<br />";

        $ages = array();
        foreach ($toalarrry as $user) {
            $banci[] = $user['banci'];
        }
        array_multisort($banci, SORT_ASC, $toalarrry); // 按照班次进行排序
        return $toalarrry;

/*        $objWriter = \PHPExcel_IOFactory::createWriter($PHPExcel, 'Excel2007');
        $objWriter->save(ROOT_PATH . "public" . DS . "uploads/excle/simple.xlsx");//文件保存路径*/
	}

	// 处理工厂的考勤表
	public function handFactory()
	{
        // Check prerequisites
        if (!file_exists(ROOT_PATH . 'public' . DS . 'uploads/excle/factorylist2.xls')) {
            exit("未找到文件：".ROOT_PATH . "public" . DS . "uploads/excle/factorylist2.xls\n");
        }

        $reader = \PHPExcel_IOFactory::createReader('Excel5'); //设置以Excel5格式(Excel97-2003工作簿)
        $PHPExcel = $reader->load(ROOT_PATH . 'public' . DS . 'uploads/excle/factorylist2.xls'); // 载入excel文件
/*        $reader = new \PHPExcel_Reader_Excel5();
        $PHPExcel = $reader ->load($file);*/

        $sheet = $PHPExcel->getSheet(0); // 读取第一個工作表
        $highestRow = $sheet->getHighestRow(); // 取得总行数
        $highestColumm = $sheet->getHighestColumn(); // 取得总列数
        ++$highestColumm;

        $columns = array("自定义编号","姓名","日期","班次","上班时间","下班时间","签到时间","签退时间","部门");
        $deleteColums = array(); // 要删除数组

//        echo "$highestColumm";

        for ($column = 'A'; $column != $highestColumm; $column++) {//列数是以A列开始
//            echo "$column";
            if(!in_array($sheet->getCell($column."1")->getValue(),$columns)){
                array_push($deleteColums,$column); // 将要删除的列添加进入一个删除数组
            }
        }

        // 要按照倒顺删除，不然会删除错误的列。
        for ($i = count($deleteColums)-1 ;$i >= 0;$i--){
//            echo "<br>";
            $znum = \PHPExcel_Cell::columnIndexFromString("Z");
            if($deleteColums[$i] < "Z"){
//                echo "删除列：".$deleteColums[$i];
                $sheet->removeColumn("".$deleteColums[$i],1);
            }
        }

        /** 循环读取每个单元格的数据,按照班次顺序排序 */
        $arrry = array();
        $toalarrry = array();
        for ($row = 2 ; $row <= $highestRow ; $row ++){
            for ($column = 'A'; $column != $highestColumm; $column++) {//列数是以A列开始
                $columnName = $sheet->getCell($column."1")->getValue();
                $columnValue = $sheet->getCell($column.$row)->getValue();
                if( $columnName == "姓名"){
                    $arrry["name"] = $columnValue;
                }
                if( $columnName == "部门"){
                    $arrry["dept"] = $columnValue;
                }
                if( $columnName == "日期"){
                    $arrry["date"] = $columnValue;
                }
                if( $columnName == "签到时间"){
                    $arrry["starttime"] = $columnValue;
                }
                if( $columnName == "签退时间"){
                    $arrry["endtime"] = $columnValue;
                }
                if( $columnName == "自定义编号"){
                    $arrry["num"] = $columnValue;
                }
            }
            array_push($toalarrry ,$arrry);
        }

//        echo "<br />";

        $ages = array();
        foreach ($toalarrry as $user) {
            $num[] = $user['num'];
        }
        array_multisort($num, SORT_ASC, $toalarrry); // 按照自定义编号进行排序

        $newarray = array();
        // 按照自定义编号删除无用的数据
       for ($m = 0 ; $m < count($toalarrry);$m++){
           $delnum = strval($toalarrry[$m]["num"]);
            if(($delnum == "40") || (strlen($delnum) == "4" && ord($delnum) != "57" || $delnum == "10033" || $delnum == "10107" )){
//                $newarray =  array_slice($toalarrry,$m,1);
                $newarray = array_merge_recursive($newarray,array_slice($toalarrry,$m,1));
            }
       }

        return $newarray;
/*        $objWriter = \PHPExcel_IOFactory::createWriter($PHPExcel, 'Excel2007');
        $objWriter->save(ROOT_PATH . "public" . DS . "uploads/excle/factsimple.xlsx");*/
	}

	// 删除指定的列
	function _remove_column($objPHPExcel,$remove_columns)
	{
		if(!$objPHPExcel
			|| !is_object ($objPHPExcel)
			|| !$remove_columns
			|| !is_array($remove_columns)
			|| count($remove_columns)<=0) return ;
    //单元格模板值,用于匹配要删除的列(在excel模板第一列)
		$cell_val = '';
    //单元格总列数
		$highestColumm = $objPHPExcel->getActiveSheet()->getHighestColumn();
		for ($column = 'A'; $column <= $highestColumm;) {
      //列数是以A列开始
			$cell_val = $objPHPExcel->getActiveSheet()->getCell($column."1");
			$cell_val = preg_replace("/[\s{}]/i","", $cell_val);
      //移出没有权限导出的列
      //移出后column不能加1,因为当前列已经移出加1后会导致删除错误的列
      //此问题浪费了几十分钟
			if(strlen($cell_val)>0 && in_array($cell_val,$remove_columns))
			{
				$objPHPExcel->getActiveSheet()->removeColumn( $column);
			}
			else
			{
				$column++;
			}
		}
}
}