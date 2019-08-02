<?php

namespace core;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use \Exception;

/**
 * Created by QiLin.
 * User: NO.01
 * Date: 2019/8/1
 * Time: 17:01
 */
class exportAction
{
    public $fileName = '导出数据demo文件';

    public $title = '测试';

    public $data = [];

    public $spreadsheet;

    public $columnLength;

    public $columnNames = [];

    public $sheet;

    public $error;

    public $startTime = '';

    const ROW_MAP = 'ABCDEFGHIJKLMNOPQRSJUVWXVZ';

    public function __construct(array $option = [])
    {
        if (isset($option['fileName'])) {
            $this->fileName = $option['fileName'];
        }
        if (isset($option['data'])) {
            $this->data = $option['data'];
        }
        if (isset($option['columnNames'])) {
            $this->columnNames = $option['columnNames'];
        }
        $this->startTime = time();
        $this->spreadsheet = new Spreadsheet();
        $this->sheet = $this->spreadsheet->getActiveSheet();
    }

    public function setTitle(string $title)
    {
        $this->title = $title;
    }

    public function getColumnLength()
    {
        return $this->columnLength = count($this->columnNames);
    }


    public function setColumnName()
    {
        foreach ($this->columnNames as $k => $name) {
            $column = self::ROW_MAP[$k] . '2';
            $this->sheet->setCellValue($column, $name);
        }
    }

    public function setColumnValue()
    {
        $init = 3;
        foreach ($this->data as $values) {
            $i = 0;
            foreach ($values as $k => $value) {
                $pCoordinate = self::ROW_MAP[$i] . $init;
                $this->sheet->setCellValue($pCoordinate, $value);
                ++$i;
            }
            ++$init;
        }
        $len = count($this->columnNames);
        $coordinate = "A2:" . self::ROW_MAP[($len - 1)] . (count($this->data) + 2);
        $this->sheet->getStyle($coordinate)->applyFromArray([
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
    }

    public function setStyle()
    {
        $this->sheet->getDefaultColumnDimension()->setWidth(30);
        //设置头部title
        $len = count($this->columnNames);
        $pRange = "A1:" . self::ROW_MAP[($len - 1)] . "1";
        $this->sheet->mergeCells($pRange);
        $this->sheet->getRowDimension("1")->setRowHeight(50);
        $this->sheet->getStyle('A1')->applyFromArray([
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
        $this->sheet->setCellValue("A1", $this->title)
            ->getStyle("A1")
            ->getFont()
            ->setSize('22')
            ->getColor()
            ->setARGB('FF0000');
        $this->sheet->setTitle($this->title);
    }

    public function check()
    {
        try {
            if (empty($this->data)) {
                throw  new Exception('data not null');
            }
            if (empty($this->columnNames)) {
                throw  new Exception('columnNames not null');
            }
            if (count($this->columnNames) > strlen(self::ROW_MAP)) {
                throw  new Exception('columnNames greater than ROW_MAP length');
            }
        } catch (Exception $exception) {
            $this->error = $exception->getMessage();
            return false;
        }
        return true;
    }

    public function export()
    {
        $this->check();
        $this->setStyle();
        $this->setColumnName();
        $this->setColumnValue();
        $writer = new Xlsx($this->spreadsheet);
        $writer->save($this->fileName . '.xlsx');
    }

    public function getError()
    {
        return $this->error;
    }

    /**
     * 断开此phpspreadsheet工作簿对象的所有工作表连接
     */
    public function __destruct()
    {
        $this->spreadsheet->disconnectWorksheets();
        unset($spreadsheet);
        $consume = time() - $this->startTime;
        echo "导出数据消耗时间：{$consume}";
    }
}