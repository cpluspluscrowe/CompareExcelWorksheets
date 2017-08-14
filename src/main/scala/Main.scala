import org.apache.poi.hssf.util.CellReference
import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.xssf.usermodel.{XSSFSheet, XSSFWorkbook}

import scala.collection.mutable.{ArrayBuffer, Map}
import java.io.File

import scala.collection.mutable

/**
  * Created by CCROWE on 8/14/2017.
  */
class Facility(){
  var map = Map[String,ArrayBuffer[String]]();
}

object Main {
  def main(args: Array[String]): Unit = {
    OpenExcel();
    println("Done!");
  }
  def GetListInDir(file:File): Array[File] ={
    if(file.exists() && file.isDirectory){
      return file.listFiles.filter(_.isDirectory);
    }
    null;
  }
  def GetExcelFileWithinFolder(file:File): Array[File] ={
    if(file.exists() && file.isDirectory){
      return file.listFiles.filter{f => f.isFile && f.getName.endsWith(".xlsx")};
    }
    null;
  }
  def GetNextConstructionActivity(currentCa:String,i:Int,sheet:XSSFSheet): String = {
    val colACellRef: CellReference = new CellReference("A" + i.toString);
    val row = sheet.getRow(i);
    if(row == null){
      return currentCa;
    }
    var cell = row.getCell(colACellRef.getCol);
    if (cell == null) {
      cell = row.createCell(colACellRef.getCol);
    }
    val formatter = new DataFormatter
    val formattedCellValue = formatter.formatCellValue(cell)
    if (formattedCellValue != null && formattedCellValue != "") {
      return formattedCellValue;
    }else{
      return currentCa;
    }
  }
  def IsTogAtRow(sheet:XSSFSheet,i:Int,togColumnLetter:String): Tuple2[Boolean,String] ={
    val colCCellRef:CellReference = new CellReference(togColumnLetter + i.toString);
    val row = sheet.getRow(colCCellRef.getRow);
    if(row == null){
      return new Tuple2(false,null);
    }
    var togCell = row.getCell(colCCellRef.getCol);
    if(togCell == null){
      togCell = row.createCell(colCCellRef.getCol);
    }
    val formatter = new DataFormatter
    val togValue = formatter.formatCellValue(togCell);
    if(togValue != null && togValue != ""){
      return new Tuple2(true,togValue);
    }
    return new Tuple2(false,togValue);
  }
  def GetTogs(workbook:XSSFWorkbook,worksheetName:String,togColumnLetter:String): Facility ={
    val facility = new Facility();
    val sheet:XSSFSheet = workbook.getSheet(worksheetName);
    var constructionActivity:String = null;
    for(i <- 2.to(sheet.getLastRowNum)){
      constructionActivity = GetNextConstructionActivity(constructionActivity,i,sheet);
      if(!facility.map.contains(constructionActivity)){
        facility.map += (constructionActivity -> new ArrayBuffer[String]())
      }
      val isTogAtRow = IsTogAtRow(sheet,i,togColumnLetter);
      if(isTogAtRow._1){
        facility.map(constructionActivity) += isTogAtRow._2;
      }
    }
    return facility;
  }
  def GetName(workbook:XSSFWorkbook): String = {
    val sheet = workbook.getSheet("Facility")
    val colACellRef: CellReference = new CellReference("A2");
    val row = sheet.getRow(colACellRef.getRow);
    var cell = row.getCell(colACellRef.getCol);
    if (cell == null) {
      cell = row.createCell(colACellRef.getCol);
    }
    val formatter = new DataFormatter
    val formattedCellValue = formatter.formatCellValue(cell)
    return formattedCellValue;
  }
  def CompareTogs(togs:Map[String,ArrayBuffer[String]],combined:Map[String,ArrayBuffer[String]]): Unit ={
    for((k,v) <- togs){
      if(!combined.contains(k)){
        println("Combined missing: " + k);
      }else{
        for(tog <- togs(k)){
          if(!combined(k).contains(tog)){
            println("Combined Ca: " + k + " missing Tog: " + tog);
          }
        }
      }
    }
  }
  def OpenExcel(): Unit ={
    val folderPath = "C:\\Create_Workbooks\\New Workbooks";
    val files:Array[File] = new File(folderPath).listFiles();
    for(excelFolder <- files){
      val excelFiles = GetExcelFileWithinFolder(excelFolder);
      for(excelFile <- excelFiles){
        val workbook:XSSFWorkbook = new XSSFWorkbook(OPCPackage.open(excelFile));
        val name = GetName(workbook);
        val togs = GetTogs(workbook,"TOGS","C");
        val drawTogs = GetTogs(workbook,"Drawings and TOGS","E");
        CompareTogs(togs.map,drawTogs.map);
        CompareTogs(drawTogs.map,togs.map);
      }
    }
  }
}
