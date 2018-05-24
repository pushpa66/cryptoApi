<?php

include("excelwriter.inc.php");
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

define("KRAKEN", "Kraken");
define("BITFINEX", "Bitfinex");
define("COINMARKETCAP", "Coincap");
define("CRYPTOCOMPARE", "CCCAGG");
define("HUBOI", "Huobi");

define("USD", "USD");
define("BTC", "BTC");

$data = getAllData();
//echo json_encode($data);
//writeData($data);

writeToSpreadSheet($data);

/**
 * @throws \PhpOffice\PhpSpreadsheet\Exception
 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
 */
function writeToSpreadSheet($data){
    $spreadsheet = new Spreadsheet();
    $sheet1 = new Worksheet($spreadsheet, 'KRAKEN');
    $sheet2 = new Worksheet($spreadsheet, 'BITFINEX');
    $sheet3 = new Worksheet($spreadsheet, 'COINMARKETCAP');
    $sheet4 = new Worksheet($spreadsheet, 'CRYPTOCOMPARE');
    $sheet5 = new Worksheet($spreadsheet, 'HUBOI');

// Data writing

    $spreadsheet->addSheet($sheet1,0);
    $spreadsheet->addSheet($sheet2,1);
    $spreadsheet->addSheet($sheet3,2);
    $spreadsheet->addSheet($sheet4,3);
    $spreadsheet->addSheet($sheet5,4);

    $sheets = array(KRAKEN => $spreadsheet->getSheet(0),
        BITFINEX => $spreadsheet->getSheet(1),
        COINMARKETCAP => $spreadsheet->getSheet(2),
        HUBOI => $spreadsheet->getSheet(3),
        CRYPTOCOMPARE => $spreadsheet->getSheet(4));

    $cellsA = array("A","B","C","D");
    $cellsB = array("E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y");
    $cellsC = array("Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT");

    $heading=array("Price","LastUpdate","LastVolume","LastVolumeTo","VolumeDay","VolumeDayTo","Volume24H","Volume24HTo",
        "OpenDay","HighDay","LowDay","Open24H","High24H","Low24H","Change24H","ChangePct24H","ChangeDay",
        "Supply","MarketCap","TotalVol24H","TotalVol24HTo");

    foreach ($sheets as $sheet){
        $sheet->getCell('A2')->setValue("Coin");
        $sheet->getStyle('A2')->getAlignment()->setWrapText(true);
        $sheet->getCell('B2')->setValue("ID");
        $sheet->getStyle('B2')->getAlignment()->setWrapText(true);
        $sheet->getCell('C2')->setValue("Name");
        $sheet->getStyle('C2')->getAlignment()->setWrapText(true);
        $sheet->getCell('D2')->setValue("Rank");
        $sheet->getStyle('D2')->getAlignment()->setWrapText(true);
    }

    foreach ($cellsB as $cell){
        foreach ($sheets as $sheet){
            $sheet->getCell($cell.'1')->setValue("USD");
            $sheet->getStyle($cell.'1')->getAlignment()->setWrapText(true);
        }
    }

    foreach ($cellsC as $cell){
        foreach ($sheets as $sheet) {
            $sheet->getCell($cell . '1')->setValue("BTC");
            $sheet->getStyle($cell . '1')->getAlignment()->setWrapText(true);
        }
    }

    for ($i = 0; $i < count($heading); $i++){
        foreach ($sheets as $sheet) {
            $sheet->getCell($cellsB[$i] . '2')->setValue($heading[$i]);
            $sheet->getStyle($cellsB[$i] . '2')->getAlignment()->setWrapText(true);
            $sheet->getCell($cellsC[$i] . '2')->setValue($heading[$i]);
            $sheet->getStyle($cellsC[$i] . '2')->getAlignment()->setWrapText(true);
        }
    }

    $count = 3;
    foreach ($data as $key => $d){
        foreach ($sheets as $sheet) {
            $sheet->getCell("A".$count)->setValue($key);
            $sheet->getStyle("A".$count)->getAlignment()->setWrapText(true);
            $sheet->getCell("B".$count)->setValue($d['Symbol']['ID']);
            $sheet->getStyle("B".$count)->getAlignment()->setWrapText(true);
            $sheet->getCell("C".$count)->setValue($d['Symbol']['Name']);
            $sheet->getStyle("C".$count)->getAlignment()->setWrapText(true);
            $sheet->getCell("D".$count)->setValue($d['Symbol']['Rank']);
            $sheet->getStyle("D".$count)->getAlignment()->setWrapText(true);
        }
//        echo json_encode($d);
        foreach ($sheets as $key => $sheet){
            $dataUSD = array();
            if(isset($d['USD'][$key])){
                $dataUSD = $d['USD'][$key];
            }
            $dataBTC = array();
            if(isset($d['BTC'][$key])){
                $dataBTC = $d['BTC'][$key];
            }

            $setDataUSD = setData($dataUSD);
            $setDataBTC = setData($dataBTC);

            for ($i = 0; $i < count($heading); $i++){
                $sheet->getCell($cellsB[$i] . $count)->setValue($setDataUSD[$i]);
                $sheet->getStyle($cellsB[$i] . $count)->getAlignment()->setWrapText(true);
                $sheet->getCell($cellsC[$i] . $count)->setValue($setDataBTC[$i]);
                $sheet->getStyle($cellsC[$i] . $count)->getAlignment()->setWrapText(true);
            }
        }

        $count++;
    }


    $sheetIndex = $spreadsheet->getIndex(
        $spreadsheet->getSheetByName('Worksheet')
    );
    $spreadsheet->removeSheetByIndex($sheetIndex);
    $writer = new Xlsx($spreadsheet);
//    $writer->save('cryptocurrencyData.xlsx');

    $filename = "cryptocurrencyData";
    header('Content-Disposition: attachment;filename="'. $filename .'.xls"'); /*-- $filename is  xsl filename ---*/
    header('Cache-Control: max-age=0');

    $writer->save('php://output');

}

function setData($data){
    $setData = array();
    $keys = array("PRICE","LASTUPDATE", "LASTVOLUME","LASTVOLUMETO","VOLUMEDAY","VOLUMEDAYTO","VOLUME24HOUR","VOLUME24HOURTO",
        "OPENDAY","HIGHDAY","LOWDAY","OPEN24HOUR","HIGH24HOUR","LOW24HOUR","CHANGE24HOUR","CHANGEPCT24HOUR","CHANGEDAY","SUPPLY",
        "MKTCAP","TOTALVOLUME24H", "TOTALVOLUME24HTO");

    foreach ($keys as $key){
        if(isset($data[$key])){
            array_push($setData, $data[$key]);
        } else {
            array_push($setData, "-");
        }
    }

    return $setData;
}

function getAllData(){
    $allData = array();

    $symbols = getSymbols();
    foreach ($symbols as $symbol) {
        $data1 = array();//USD
        $data2 = array();//BTC
        $data3 = array();
        if($symbol['Symbol'] == USD){
            $data2 = getData($symbol['Symbol'], BTC);
        } else if ($symbol['Symbol'] == BTC){
            $data1 = getData($symbol['Symbol'], USD);
        } else {
            $data1 = getData($symbol['Symbol'], USD);
            $data2 = getData($symbol['Symbol'], BTC);
        }

        $data3["USD"] = $data1;
        $data3["BTC"] = $data2;
        $data3["Symbol"] = $symbol;
        $allData[$symbol['Symbol']] = $data3;
    }
    return $allData;
}

function getSymbols(){
    $symbols = array();
    $symbolCount = 50;

    $curl = curl_init();

    curl_setopt_array($curl, array(
        CURLOPT_URL => "https://api.coinmarketcap.com/v2/ticker/?limit=$symbolCount",
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_ENCODING => "",
        CURLOPT_MAXREDIRS => 10,
        CURLOPT_TIMEOUT => 30,
        CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
        CURLOPT_CUSTOMREQUEST => "GET",
        CURLOPT_HTTPHEADER => array(
            "Cache-Control: no-cache",
            "Postman-Token: 343dd779-07f0-497a-9abc-d3094b078606"
        ),
    ));

    $response = curl_exec($curl);
    $err = curl_error($curl);

    curl_close($curl);

    if ($err) {
        echo "cURL Error #:" . $err;
    } else {
        $response = json_decode($response, true);
        $symbolsData = $response['data'];
        foreach($symbolsData as $symbolData){
            $symbol = array();
            $symbol['ID'] = $symbolData['id'];
            $symbol['Name'] = $symbolData['name'];
            $symbol['Rank'] = $symbolData['rank'];
            $symbol['Symbol'] = $symbolData['symbol'];
            array_push($symbols, $symbol);
        }

    }

//    var_dump($symbols);
    return $symbols;
    
}

function getData($fromSymbol, $toSymbol){

    $dataSet = array();
    $curl = curl_init();

    curl_setopt_array($curl, array(
        CURLOPT_URL => "https://min-api.cryptocompare.com/data/top/exchanges/full?fsym=$fromSymbol&tsym=$toSymbol&limit=200",
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_ENCODING => "",
        CURLOPT_MAXREDIRS => 10,
        CURLOPT_TIMEOUT => 30,
        CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
        CURLOPT_CUSTOMREQUEST => "GET",
        CURLOPT_HTTPHEADER => array(
        "Cache-Control: no-cache",
        "Postman-Token: 1b49906c-b7a2-4c67-aed4-b1e7793f1c26"
        ),
    ));
    
    $response = curl_exec($curl);
    $err = curl_error($curl);
    
    curl_close($curl);
    
    if ($err) {
        echo "cURL Error #:" . $err;
    } else {
        $response = json_decode($response, true);

        $readData =  array();
        $dataSet[KRAKEN] = array();
        $dataSet[BITFINEX] = array();
        $dataSet[COINMARKETCAP] = array();
        $dataSet[HUBOI] = array();
        $dataSet[CRYPTOCOMPARE] = array();


        if(isset($response['Data']['Exchanges'])){
            $readData = $response['Data']['Exchanges'];
        }

        if(isset($response['Data']['AggregatedData'])){
            $dataSet[CRYPTOCOMPARE] = $response['Data']['AggregatedData'];
        }

        $count = 0;
        foreach ($readData as $data) {
            if($count == 4){
                break;
            }
            if ( $data['MARKET'] == KRAKEN){
                $count ++;
                $dataSet[KRAKEN] = $data;
            } else if ($data['MARKET'] == BITFINEX){
                $count ++;
                $dataSet[BITFINEX] = $data;
            }else if ($data['MARKET'] == COINMARKETCAP){
                $count ++;
                $dataSet[COINMARKETCAP] = $data;
            }else if ($data['MARKET'] == HUBOI){
                $count ++;
                $dataSet[HUBOI] = $data;
            }
        }
    }

    return $dataSet;
}
