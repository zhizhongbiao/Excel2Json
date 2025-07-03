package com.zzb.excel2json

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import com.fasterxml.jackson.module.kotlin.*
import com.fasterxml.jackson.databind.ObjectMapper
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.VerticalAlignment

fun main() {
    val inputDir = "json"
    val outputExcel = "language${File.separator}language.xlsx"
    jsonToExcel(inputDir, outputExcel)
}

fun jsonToExcel(inputDir: String, outputExcel: String) {
    val mapper = ObjectMapper().registerKotlinModule()
    val langFiles = File(inputDir).listFiles { f -> f.extension == "json" && f.name != "remarks.json" } ?: return

    val langMaps = mutableMapOf<String, Map<String, String>>()
    langFiles.forEach { file ->
        val lang = file.nameWithoutExtension
        val map: Map<String, String> = mapper.readValue(file)
        langMaps[lang] = map
    }
    val allKeys = langMaps.values.flatMap { it.keys }.toSet().sorted()
    val langs = langMaps.keys.sorted()

    // 读取备注
    val remarkFile = File(inputDir, "remarks.json")
    val remarkMap: Map<String, String> = if (remarkFile.exists()) mapper.readValue(remarkFile) else langs.associateWith { "" }

    // Excel workbook和样式
    val workbook = XSSFWorkbook()
    val sheet = workbook.createSheet("Sheet1")

    // 创建加粗且居中的样式（用于备注行和语种标识行）
    val boldCenterStyle = workbook.createCellStyle().apply {
        setFont(workbook.createFont().apply { bold = true })
        alignment = HorizontalAlignment.CENTER
        verticalAlignment = VerticalAlignment.CENTER
    }
    // 创建普通居中样式（用于普通数据）
    val centerStyle = workbook.createCellStyle().apply {
        alignment = HorizontalAlignment.CENTER
        verticalAlignment = VerticalAlignment.CENTER
    }

    // 写备注行（加粗+居中）
    val remarkRow = sheet.createRow(0)
    val remarkCell0 = remarkRow.createCell(0)
    remarkCell0.setCellValue("备注")
    remarkCell0.cellStyle = boldCenterStyle
    langs.forEachIndexed { i, lang ->
        val cell = remarkRow.createCell(i + 1)
        cell.setCellValue(remarkMap[lang] ?: "")
        cell.cellStyle = boldCenterStyle
    }

    // 2. 语种标识行样式（加粗+居中+红色字体）
    val boldRedFont = workbook.createFont().apply {
        bold = true
        color = IndexedColors.RED.index
    }
    val boldRedCenterStyle = workbook.createCellStyle().apply {
        setFont(boldRedFont)
        alignment = HorizontalAlignment.CENTER
        verticalAlignment = VerticalAlignment.CENTER
    }

    // 写语种代码行（加粗+居中+红色字体）
    val codeRow = sheet.createRow(1)
    val codeCell0 = codeRow.createCell(0)
    codeCell0.setCellValue("key")
    codeCell0.cellStyle = boldRedCenterStyle
    langs.forEachIndexed { i, lang ->
        val cell = codeRow.createCell(i + 1)
        cell.setCellValue(lang)
        cell.cellStyle = boldRedCenterStyle
    }

    // 写数据行（全部居中）
    allKeys.forEachIndexed { rowIdx, key ->
        val row = sheet.createRow(rowIdx + 2)
        val keyCell = row.createCell(0)
        keyCell.setCellValue(key)
        keyCell.cellStyle = centerStyle
        langs.forEachIndexed { i, lang ->
            val value = langMaps[lang]?.get(key) ?: ""
            val cell = row.createCell(i + 1)
            cell.setCellValue(value)
            cell.cellStyle = centerStyle
        }
    }

    // 自动列宽
    (0..langs.size).forEach { sheet.autoSizeColumn(it) }

    File(outputExcel).outputStream().use { workbook.write(it) }
    println("Merged Excel generated: $outputExcel")
}