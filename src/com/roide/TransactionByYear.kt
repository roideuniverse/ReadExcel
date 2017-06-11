package com.roide

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.text.SimpleDateFormat
import java.util.*
import kotlin.collections.ArrayList

/**
 * Created by ks on 6/10/17.
 */

private val dateFormatter = SimpleDateFormat("dd/MM/yyyy")

fun main(args: Array<String>) {
   val filePath = "/Users/ks/Desktop/files/CompanyName_RPT500HC.xlsx"

  val file = File(filePath)
  val dirName = file.parentFile
  val fileName = file.name

  val excelFileInputStream = FileInputStream(file)
  val readWorkbook = XSSFWorkbook(excelFileInputStream)
  val companyNames = readCompanyNames(readWorkbook.getSheet("Sheet1"))

  val writeWordbook = XSSFWorkbook()

  val sheetIterator = readWorkbook.sheetIterator()
  while(sheetIterator.hasNext()) {
    val sheet = sheetIterator.next()
    System.out.println("name=" + sheet.sheetName)
    if (sheet.sheetName == "Sheet1") {
       continue
    }
    val entries = read(sheet)
    val mapEntries = createNewEntries(entries, companyNames)
    val newSheet = writeWordbook.createSheet(sheet.sheetName)
    writeToSheet(newSheet, mapEntries, companyNames)
  }

  autoSizeColumns(writeWordbook)
  // Write to New ExcelSheet
  val newFileAbsolutePath = "" + dirName.absolutePath + File.separator + "New_" + fileName
  val newFile = File(newFileAbsolutePath)
  val outputStream = FileOutputStream(newFile)
  writeWordbook.write(outputStream)
  writeWordbook.close()
  readWorkbook.close()

}

fun writeToSheet(sheet: Sheet, newEntries: Map<String, NewEntry>, cNames: List<String>) {
  var rowNum = 0
  for (cname in cNames) {
    val row = sheet.createRow(rowNum++)
    for (i in 0..10) {
      row.createCell(i)
    }
    row.getCell(0).setCellValue(cname)
    val newEntry = newEntries.get(cname)
    if (newEntry != null) {
      row.getCell(1).setCellValue(newEntry.year04Sum)
      row.getCell(2).setCellValue(newEntry.year05Sum)
      row.getCell(3).setCellValue(newEntry.year06Sum)
      row.getCell(4).setCellValue(newEntry.year07Sum)
      row.getCell(5).setCellValue(newEntry.year08Sum)
      row.getCell(6).setCellValue(newEntry.year09Sum)
      row.getCell(7).setCellValue(newEntry.year10Sum)
      row.getCell(8).setCellValue(newEntry.year11Sum)
    }
  }
}

fun createNewEntries(entries: List<Entry>, companyNames: List<String>): Map<String, NewEntry> {
  val mapEntries = TreeMap<String, NewEntry>()
  for (entry in entries) {
    if (mapEntries.containsKey(entry.companyName)) {
      val newEntry = addYearSum(entry, mapEntries.get(entry.companyName))
      mapEntries.put(entry.companyName, newEntry)
    } else {
      val newEntry = addYearSum(entry, null)
      mapEntries.put(entry.companyName, newEntry)
    }
  }
  return mapEntries
}

fun addYearSum(entry: Entry, newEntry: NewEntry?) : NewEntry {
  val companyName = entry.companyName
  var s04 = newEntry?.year04Sum ?: 0.0
  var s05 = newEntry?.year05Sum ?: 0.0
  var s06 = newEntry?.year06Sum ?: 0.0
  var s07 = newEntry?.year07Sum ?: 0.0
  var s08 = newEntry?.year08Sum ?: 0.0
  var s09 = newEntry?.year09Sum ?: 0.0
  var s10 = newEntry?.year10Sum ?: 0.0
  var s11 = newEntry?.year11Sum ?: 0.0

  val year = (entry.date.split("/")[2]).toInt()
  when(year) {
    2004 -> s04 += entry.transactionValue
    2005 -> s05 += entry.transactionValue
    2006 -> s06 += entry.transactionValue
    2007 -> s07 += entry.transactionValue
    2008 -> s08 += entry.transactionValue
    2009 -> s09 += entry.transactionValue
    2010 -> s10 += entry.transactionValue
    2011 -> s11 += entry.transactionValue
  }
  return NewEntry(companyName, s04, s05, s06, s07, s08, s09, s10, s11)
}

fun readCompanyNames(sheet :Sheet): List<String> {
  val list = ArrayList<String>()
  val rowIterator = sheet.rowIterator()
  while(rowIterator.hasNext()) {
    val row = rowIterator.next()
    val cName = row.getCell(0).stringCellValue
    list.add(cName)
   // println(cName)
  }
  return list
}

fun read(sheet :Sheet): List<Entry> {
  val iterator = sheet.rowIterator()
  //var prevCompanyName = ""
  //var prevDate : String = ""
  //var prevPartyType : String = ""
  val entryList = ArrayList<Entry>()
  while (iterator.hasNext()) {
    val row = iterator.next()
    if (row.getCell(0) == null)
      continue
    val companyName = row.getCell(0).stringCellValue
    //prevCompanyName = companyName

    val tempVal = row.getCell(1)
    //val type = tempVal?.cellTypeEnum
    val entryDate = tempVal.stringCellValue
    //if (type == CellType.NUMERIC) dateFormatter.format(row.getCell(1).dateCellValue) else prevDate
    //prevDate = entryDate

    val partyType = row.getCell(2).stringCellValue
    //prevPartyType = partyType

    val partyName = row.getCell(3).stringCellValue
    val transactionValue = row.getCell(6).numericCellValue

    val expression = row.getCell(4).stringCellValue
    val transactionType = row.getCell(5).stringCellValue
    val unitVal = row.getCell(7).stringCellValue

    val entry = Entry(companyName, entryDate, partyType, partyName, transactionValue, expression, transactionType, unitVal)
    entryList.add(entry)
    //println(entry)
  }
  return entryList
}

data class NewEntry(val companyName:String,
                    val year04Sum: Double,
                    val year05Sum: Double,
                    val year06Sum: Double,
                    val year07Sum: Double,
                    val year08Sum: Double,
                    val year09Sum: Double,
                    val year10Sum: Double,
                    val year11Sum: Double)
