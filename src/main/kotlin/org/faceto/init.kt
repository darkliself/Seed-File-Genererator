package org.faceto

import com.github.doyaaaaaken.kotlincsv.dsl.csvReader
import com.monitorjbl.xlsx.StreamingReader
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream


fun main() {
    val dublicate = arrayListOf(
        "Binder Accessories",
        "Sports & Recreation Books",
        "Geriatric Seating",
        "Blank Media",
        "Card Readers & Adapters"
    )
    val categories = arrayListOf("Sports & Recreation Books", "Geriatric Seating", "Card Readers & Adapters")
    val listOfCategoriesRows = getListWithCategories(categories)
}


fun getListWithCategories(cat: ArrayList<String>) {
    val listOfCategories = mutableListOf<MutableList<String>>()
    val staplesDumpFile = FileInputStream(File("./full_dump.xlsx"))
    val workBook = StreamingReader.builder()
        .rowCacheSize(100) // number of rows to keep in memory (defaults to 10)
        .bufferSize(4096) // buffer size to use when reading InputStream to file (defaults to 1024)
        .open(staplesDumpFile) // Input

    workBook.getSheetAt(0).forEach { row ->
        cat.forEach { name ->
            if (row.getCell(3).stringCellValue == name) {
                val tmp = mutableListOf<String>()
                row.forEach { cell ->
                    tmp.add(cell.stringCellValue)
                }
                listOfCategories.add(tmp)
            }
        }
    }
    generateTable(listOfCategories)
}

fun generateTable(listOfCategories: MutableList<MutableList<String>>) {
    val tt = mutableMapOf<String, MutableList<String>>()
    var tmp = ""
    tt["filter"] = mutableListOf(
        "{L=0}", "{I=0}", "NAME", "EXPRESSION",	"EXPRESSION_STATUS", "#8", "#9", "#12",	"#45", "#10", "#13", "#14", "#48")

    listOfCategories.forEach { row ->
        tmp = ""
        if (row[10] == "List Of Values") {

            if (tt.get(row[3] + row[9]) == null) {
                tmp = row[14]
            } else {
                tmp = tt.get(row[3] + row[9])!!.get(10) + "|" + row[14]
            }

        }
        // create seed table for category
        tt.put(row[3] + row[9], mutableListOf(
            row[3],
            row[4],
            row[8],
            "Add(R(\"cnet_common_${row[9]}\"));",
            "ready",
            row[13].replace("N", "OPTIONAL").replace("Y", "REQUIRED"),
            row[10].replace("number", "DECIMAL")
                .replace("text", "TEXT")
                .replace("List Of Values", "LIST"),
            row[9],
            "1",
            "0",
            tmp,
            row[16],
            if (tmp != "") (tmp.count { it == '|' } + 1).toString() else ""
        ))

    }
    addCommonSectionItems(tt)
    writeToFile(tt)
}

fun addCommonSectionItems(table: MutableMap<String, MutableList<String>>) {
    val listOfCommonItems = getCSVColumnsByIndex("./COMMON.csv", 47) //////////////////////////////
    val tmpMap = mutableMapOf<String, MutableList<String>>()

    for ((k, v) in table) {
        if (!listOfCommonItems.contains("cnet_common_${v[7]}")) {
            tmpMap.put("COMMON_${v[7]}", mutableListOf(
                "COMMON",
                "COMMON_SECTIONS",
                v[2],
                "",
                "",
                "OPTIONAL",
                v[6],
                "cnet_common_${v[7]}",
                v[8],
                v[9],
                v[10],
                v[11],
                v[12]
            ))

        }
    }
    for ((key, value) in tmpMap) {
        table[key] = value
    }
}

fun getCSVColumnsByIndex(filepath: String, indexOfColumns: Int): MutableSet<String> {
    val list = mutableSetOf<String>()
    csvReader().open(filepath) {
        readAllAsSequence().forEach { row ->
            list.add(row[indexOfColumns])
        }
    }
    return list
}

fun writeToFile(rowsToWrite: MutableMap<String, MutableList<String>>) {
    val myWorkBook = XSSFWorkbook()
    val myWorkList = myWorkBook.createSheet("seed_request")
    var row = 1
    var column = 0
    myWorkList.createRow(0)

    rowsToWrite.forEach { (key, value) ->
        myWorkList.createRow(row)
        value.forEach {
            myWorkList.getRow(row).createCell(column).setCellValue(it)
            column++
        }
        row++
        column = 0
    }
    println(myWorkBook)
    val output = FileOutputStream("./seed_test.xlsx")
    myWorkBook.write(output)
}