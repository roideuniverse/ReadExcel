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
 * Created by ks on 6/4/17.
 */

private val FIRST_ROW_WITH_DATA = 3
private val dateFormatter = SimpleDateFormat("dd/MM/yyyy")
private val DEBUG = false

fun main(args: Array<String>) {

  var filePath : String = ""
  var filterPartyName: String = ""
  if (!DEBUG) {
    val parsedArgs = parseInput(args) ?: return
    filePath = parsedArgs.first
    filterPartyName = parsedArgs.second
  } else {
    filePath = "/Users/ks/Desktop/files/RPT500EKPC.xlsx"
    filterPartyName = "All Ent. over which KMP have control/SI"
  }

  val file = File(filePath)
  val dirName = file.parentFile
  val fileName = file.name

  val excelFileInputStream = FileInputStream(file)
  val readWorkbook = XSSFWorkbook(excelFileInputStream)

  val writeWordbook = XSSFWorkbook()

  val sheetIterator = readWorkbook.sheetIterator()
  while (sheetIterator.hasNext()) {
    val sheet = sheetIterator.next()
    //println(sheet.sheetName)

    val sheetReader = SheetReader(sheet)
    val entries = sheetReader.read().filter {
      it.partyName == filterPartyName
    }
    val result = query(entries)
    // println(result.size)
    val newSheet = writeWordbook.createSheet(sheet.sheetName)
    writeToSheet(newSheet, result)
  }

  // Auto Size column
  autoSizeColumns(writeWordbook)

  // Write to New ExcelSheet
  val newFileAbsolutePath = "" + dirName.absolutePath + File.separator + "New_" + fileName
  val newFile = File(newFileAbsolutePath)
  val outputStream = FileOutputStream(newFile)
  writeWordbook.write(outputStream)
  writeWordbook.close()
  readWorkbook.close()
}

fun parseInput(args: Array<String>): Pair<String, String>? {
  var pair: Pair<String, String>? = null
  loop@ for (arg in args) {
    if (!arg.startsWith("--") || !arg.contains("=")) break
    val argSplit = arg.split("=")
    val key = argSplit[0].removePrefix("--")
    val value = argSplit[1].trim()
    when (key) {
      "file" -> {
        val file = File(value)
        if (!file.exists()) {
          println("File Does not exist")
          break@loop
        }
        pair = Pair(value, "")
      }
      "filterPartyName" -> {
        pair = Pair(pair!!.first, value)
        return pair
      }
    }
  }
  println("-------------------------------")
  println("Unexpected arguments to Program")
  println("try:")
  println("java -jar ReadExcel.jar --file=\"path to file\" --filterPartyName=\"All Holding Company\"")
  println("-------------------------------")
  return null
}

fun query(entries: List<Entry>): Map<Triple<String, String, String>, Pair<Entry, Double>> {
  //val prim = Triple<String, String, String>()
  val map = LinkedHashMap<Triple<String, String, String>, Pair<Entry, Double>>()
  for (entry in entries) {
    val key = Triple(entry.companyName, entry.date, entry.partyType)
    if (map.containsKey(key)) {
      val valPair = map.get(key)!!
      val sum = valPair.second + entry.transactionValue
      map.put(key, Pair(entry, sum))
    } else {
      map.put(key, Pair(entry, entry.transactionValue))
    }
  }
  return map
}

fun writeToSheet(sheet: Sheet, data: Map<Triple<String, String, String>, Pair<Entry, Double>>) {
  var rowNum = 0
  var prevCompany = ""
  for (pair in data.values) {
    val entry = pair.first

    // Add an empty row if there is no company
    if (prevCompany != entry.companyName) {
      if (prevCompany != "") {
        sheet.createRow(rowNum++).createCell(0)
      }
      prevCompany = entry.companyName
    }

    val sum = pair.second
    val row = sheet.createRow(rowNum++)
    for (i in 0..7) {
      val cell = row.createCell(i)
      when (i) {
        0 -> cell.setCellValue(entry.companyName)
        1 -> cell.setCellValue(entry.date)
        2 -> cell.setCellValue(entry.partyType)
        3 -> cell.setCellValue(entry.partyName)
        4 -> cell.setCellValue(entry.expression)
        5 -> cell.setCellValue(entry.transactionType)
        6 -> cell.setCellValue(sum)
        7 -> cell.setCellValue(entry.valueUnit)
      }
    }
  }
}

class SheetReader(excelSheet: Sheet) {
  val sheet = excelSheet

  fun read(): List<Entry> {
    val iterator = sheet.rowIterator()
    for (i in 0..FIRST_ROW_WITH_DATA) {
      if (iterator.hasNext()) {
        iterator.next()
      } else {
        return Collections.emptyList()
      }
    }

    var prevCompanyName = ""
    var prevDate : String = ""
    var prevPartyType : String = ""
    val entryList = ArrayList<Entry>()
    while (iterator.hasNext()) {
      val row = iterator.next()
      val companyName = if (row.getCell(0) != null) row.getCell(0).stringCellValue else prevCompanyName
      prevCompanyName = companyName

      val tempVal = row.getCell(1)
      val type = tempVal?.cellTypeEnum
      val entryDate = if (type == CellType.NUMERIC) dateFormatter.format(row.getCell(1).dateCellValue) else prevDate
      prevDate = entryDate

      val partyType = if (row.getCell(3) != null) row.getCell(3).stringCellValue else prevPartyType
      prevPartyType = partyType

      val partyName = if (row.getCell(4) != null) row.getCell(4).stringCellValue else continue
      val transactionValue = row.getCell(7).numericCellValue

      val expression = row.getCell(5).stringCellValue
      val transactionType = row.getCell(6).stringCellValue
      val unitVal = row.getCell(8).stringCellValue

      val entry = Entry(companyName, entryDate, partyType, partyName, transactionValue, expression, transactionType, unitVal)
      entryList.add(entry)
      //println(entry)
    }
    return entryList
  }
}

fun autoSizeColumns(workbook: XSSFWorkbook) {
  val numberOfSheets = workbook.numberOfSheets
  for (i in 0..numberOfSheets - 1) {
    val sheet = workbook.getSheetAt(i)
    if (sheet.physicalNumberOfRows > 0) {
      val row = sheet.getRow(0)
      val cellIterator = row.cellIterator()
      while (cellIterator.hasNext()) {
        val cell = cellIterator.next()
        val columnIndex = cell.columnIndex
        sheet.autoSizeColumn(columnIndex)
      }
    }
  }
}

data class Entry(val companyName: String,
                 val date: String,
                 val partyType: String,
                 val partyName: String,
                 val transactionValue: Double,
                 val expression: String,
                 val transactionType: String,
                 val valueUnit: String)

