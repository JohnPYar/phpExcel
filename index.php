<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$url = 'https://opt.wspitaly.com.ua/remains/Price%20WSP%20Italy.xls';

$fileNewCatalog = 'originalWspItaly.xls';

copy($url, $fileNewCatalog);

//$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($fileNewCatalog);


$worksheet = $spreadsheet->getActiveSheet();

$maxRow = $worksheet->getHighestRow();

echo $maxRow;
// НАСТРОЙКИ для редактирования

// массив замены значений столбца Остатки
$stocksReplaceValues = [
   'під замовлення' => +'0',
   '*' => +'0'
];

// недостающий производитель
$manufacturer = 'GREEN line';

// тексты Описаний для каждого Производителя
$descriptions = [
   'GREEN line' => 'В комплектацию поставки входят:

крышка;
крепление;
логотип;
Колёсные литые диски WSP Green line изготовлены при помощи качественного современного оборудования с применением технологии литья под низким давлением - LOW PRESSURE.
',
   'MAK' => 'Итальянская компания Mak занимается производством автомобильных деталей с 90-х годов. Диски Mak , разработанные для сегмента класса «премиум», отвечают самым жестким эксплуатационным требованиям и высочайшим международным стандартам качества. Теперь диски Мак известны не только в Италии, но и во всем мире. Многие автолюбители оценили уже их великолепный дизайн, высокое качество и надежность. 
Бренд дисков Mak уверенно занимает лидирующие позиции в своей отрасли, чему способствуют технологические инновации, внедряемые в производственный процесс. 
• Прочность и долговечность, износоустойчивость. Даже при длительной интенсивной эксплуатации автомобильные диски Мак сохраняют внешний вид. 
• Обеспечивают хорошее сцепление, управляемость и маневренность автомобиля, стабильность на любом дорожном покрытии. 
• Диски Mak влияют на экономный расход топлива, длительность работы тормозной системы, безопасность вождения.
',
   'WSP Italy' => 'Автомобильные литые диски WSP Italy разрабатываются и тестируются в полном соответствии с самыми строгими европейскими государственными и международными директивами и стандартами. Немецкий стандарт TÜV, Японские стандарты JWL-VIA, UN/ECE R124. Официально аккредитованы: Министерством транспорта Италии и VCA –Великобритания. 
Важным фактором является наличие сертификата омологации, который обязателен для запасных частей автомобиля в странах Европейского Союза. 
Итальянский производитель, применяя передовые технологии в производстве, добился высоких эксплуатационных показателей: 
- увеличена нагрузка минимум на 25% (диски более прочные),
- полное соответствие параметрам производителя авто (потребитель не лишается гарантии), 
- более устойчивая покраска в сравнении с дешевыми аналогами (внешний вид не утратит свой блеск через годы). 
ТМ WSP Italy – это лидер в сегменте дисков "Replica"
'
];

// НАСТРОЙКИ end

// меняем в столбце Остатки некорректные значения на 0

// получаем значения ячеек столбца "Остатки"
$stocksColumn = $worksheet -> rangeToArray(
   'N3:N'.$maxRow,
   NULL,
   TRUE,
   FALSE,
   FALSE
);
//echo '<pre>';
//var_dump($stocksColumn);
//echo '</pre>';


// меняем значения в столбце по условиям из массива замены
foreach ($stocksColumn as $cell => $value){
   foreach ($value as $key => $item){
      foreach ($stocksReplaceValues as $replace => $to){
         if ($item == $replace){
            $stocksColumn[$cell][$key] = $to;
         }
      }
   }
}
//echo '<pre>';
//var_dump($stocksColumn);
//echo '</pre>';

// вставляем столбец с измененными значениями обратно
$worksheet -> fromArray(
  $stocksColumn,
  NULL,
  'N3',
   TRUE
);

// добавляем в столбце Производитель Green Line

$autoModelColumn = $worksheet->rangeToArray(
   'D3:D'.$maxRow,
   NULL,
   TRUE,
   FALSE,
   FALSE
);
//echo '<pre>';
//var_dump($autoModelColumn);
//echo '</pre>';

foreach ($autoModelColumn as $cell => $value){
   foreach ($value as $key => $item){
         if ($item == $manufacturer){
            $worksheet->setCellValue('A'. ($cell+3), $manufacturer);
         }
   }
}

// добавляем колонку с Описанием
$descriptionColumn = 'описание';
$worksheet -> setCellValue('R2', $descriptionColumn);


// добавляем тексты Описаний для каждого Производителя
$manufacturerColumn = $worksheet->rangeToArray(
   'A3:A'.$maxRow,
   NULL,
   TRUE,
   FALSE,
   FALSE
);

foreach ($manufacturerColumn as $cell => $value){
   foreach ($value as $key => $item){
      switch ($item){
         case "GREEN line":
            $worksheet->setCellValue('R' . ($cell+3), $descriptions['GREEN line']);
            break;
         case "MAK":
            $worksheet->setCellValue('R' . ($cell+3), $descriptions['MAK']);
            break;
         case "WSP Italy":
            $worksheet->setCellValue('R' . ($cell+3), $descriptions['WSP Italy']);
            break;
      }
   }
}

// сохраняем измененные данные в новый файл excel
$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xls($spreadsheet);
$writer->save('newWspItaly.xls');

echo 'File is ready!';

//очищаем память
$spreadsheet->disconnectWorksheets();
unset($spreadsheet);