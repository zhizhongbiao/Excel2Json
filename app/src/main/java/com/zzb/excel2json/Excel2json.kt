package com.zzb.excel2json

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import com.fasterxml.jackson.module.kotlin.*
import com.fasterxml.jackson.databind.ObjectMapper

fun main() {
    val excelPath = "language${File.separator}language.xlsx"
    val outputDir = "json"
    excelToJson(excelPath, outputDir)
}

fun excelToJson(excelPath: String, outputDir: String) {
    val file = File(excelPath)
    val workbook = XSSFWorkbook(file.inputStream())
    val sheet = workbook.getSheetAt(0)
    val headerRow = sheet.getRow(0)
    val langCodes = mutableListOf<String>()
    val remarks = mutableListOf<String>()
    // 读取备注行（第一行）和语言代码（第二行）
    val langCount = headerRow.lastCellNum
    for (col in 1 until langCount) {
        val remark = headerRow.getCell(col)?.stringCellValue?.trim() ?: ""
        remarks.add(remark)
    }
    val langRow = sheet.getRow(1)
    for (col in 1 until langCount) {
        val code = langRow.getCell(col)?.stringCellValue?.trim() ?: ""
        if (code.isNotEmpty()) langCodes.add(code)
    }
    // 读取实际内容，第三行起
    val langMaps = langCodes.associateWith { mutableMapOf<String, String>() }
    for (rowIdx in 2..sheet.lastRowNum) {
        val row = sheet.getRow(rowIdx) ?: continue
        val key = row.getCell(0)?.stringCellValue?.trim() ?: continue
        if (key.isEmpty()) continue
        for ((langIdx, lang) in langCodes.withIndex()) {
            val value = row.getCell(langIdx + 1)?.stringCellValue?.trim() ?: ""
            langMaps[lang]?.put(key, value)
        }
    }
    // 输出json
    File(outputDir).mkdirs()
    val mapper = ObjectMapper().registerKotlinModule().writerWithDefaultPrettyPrinter()
    langMaps.forEach { (lang, map) ->
        val jsonFile = File(outputDir, "$lang.json")
        mapper.writeValue(jsonFile, map)
        println("Generated: ${jsonFile.absolutePath}")
    }
    // 输出备注，key=langCode，value=备注
    val remarkMap = langCodes.zip(remarks).toMap()
    val remarkFile = File(outputDir, "remarks.json")
    mapper.writeValue(remarkFile, remarkMap)
    println("Generated: ${remarkFile.absolutePath}")
}