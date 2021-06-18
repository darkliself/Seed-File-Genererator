package org.faceto

import com.github.doyaaaaaken.kotlincsv.dsl.csvReader
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream


fun main(args: Array<String>) {
    println("Hello, World")


//    val myWorkBook = XSSFWorkbook()
//    val myWorkList =  myWorkBook.createSheet("test sheet")
//    myWorkList.createRow(0).createCell(1).setCellValue("here we go!")
//    val output = FileOutputStream("./seed_result.xlsx")
//    myWorkBook.write(output)
    val myWorkBook2 = XSSFWorkbook()
    //val input = FileInputStream("./request.xlsx")
    readFromExcelFile("./seed_result.xlsx")

}


fun readFromExcelFile(filepath: String) {
    val listOfValues = mutableMapOf<String, String>()
    val inputStream = FileInputStream(filepath)
    //Instantiate Excel workbook using existing file:
    val xlWb = WorkbookFactory.create(inputStream)
    val xlWs = xlWb.getSheetAt(0)
    val someString = mutableMapOf<String, String>()
    val listOfCommonIds = getListOfCommonIdentifiers("./COMMON.csv")
    val listOfCategories = arrayOf("Air Purifiers", "Projector Screens", "Padlocks", "Microwaves", "Dehumidifiers")
    val listOfValuesForCommon = mutableMapOf<String, String>()
    listOfCategories.forEach {
        generateCellsForCategory(xlWs, someString,  it, listOfCommonIds, listOfValuesForCommon)
        getListOfValuesForCategory(xlWs, listOfValues, it)
        someString.remove("${it}Attribute_Name")
        someString.remove("${it}Next Day Delivery")
    }


    // remove not needed patterns
    for (elem in listOfValues) {
        if (someString[elem.key] !== null) {
            someString[elem.key] = someString[elem.key].toString()
                .replace("pattern for values", elem.value)
                .replace("pattern for count", (elem.value.count{it == '|'} + 1).toString())
        }
    }

    writeToFile(someString)
    println(someString)
}

fun generateCommonItems( listOfItemsId : MutableMap<String, List<String>>, listOfCommonIds : MutableSet<String>, result: MutableMap<String, String>, listOfValuesForCommon: MutableMap<String, String>) {

    var hasValue = false
    for (elem in listOfItemsId) {
        if (!listOfCommonIds.contains("cnet_common_${elem.key}")) {
            result[elem.key] = ("COMMON"
                    + "<>" + "COMMON_SECTIONS"
                    + "<>" + elem.value[1]
                    + "<>"
                    + "<>"
                    + "<>" + "OPTIONAL"
                    + "<>" + elem.value[2]
                    + "<>" + "cnet_common_${elem.value[0]}"
                    + "<>" + "1"
                    + "<>" + "0"
                    + "<>" + listOfValuesForCommon[elem.key].toString().replace("null", "")
                    + "<>" + elem.value[4]
                    )
            if (listOfValuesForCommon[elem.key].toString().replace("null", "") !== "") {
                result[elem.key] = result[elem.key] + "<>" + (listOfValuesForCommon[elem.key].toString().count{it == '|'} + 1)
            }
            println(elem)
        }

    }


}

fun getListOfCommonIdentifiers(filepath: String):  MutableSet<String> {
    var list = mutableSetOf<String>()
    csvReader().open(filepath) {
        readAllAsSequence().forEach { row ->
            list.add(row[47])
        }
    }
    println(list.count())
   return list
}

fun generateCellsForCategory(sheet: Sheet, result: MutableMap<String, String>, categoryName: String, listOfCommonIds: MutableSet<String>, listOfValuesForCommon: MutableMap<String, String>) {
    var shortName = ""
    var catKey = ""
    var listOfItemsId = mutableMapOf<String, List<String>>()
    var listOfValuesForCommon = mutableMapOf<String, String>()
    for (elem in sheet) {
            if (elem.getCell(3).toString() == categoryName) {
                if (elem.getCell(14).toString() !== "") {
                    if (listOfValuesForCommon[elem.getCell(9).toString()] == null) {
                        listOfValuesForCommon[elem.getCell(9).toString()] = elem.getCell(14).toString()
                    } else {
                        listOfValuesForCommon[elem.getCell(9).toString()] =
                            listOfValuesForCommon[elem.getCell(9).toString()] + "|" + elem.getCell(14).toString()
                    }
                }
            }

    }
    listOfValuesForCommon.forEach { t, u ->
        println(t + " " +u)
    }

    // add SP_ items for this cat for generating common items
    for (elem in sheet) {
        if (result[elem.getCell(8).toString()] !== "") {
            var comment = ""
            if (elem.getCell(16) !== null) {
                comment = elem.getCell(16).toString()
            }
            if (categoryName == elem.getCell(3).toString()) {
                listOfItemsId[elem.getCell(9).toString()] = listOf(
                    elem.getCell(9).toString(),
                    elem.getCell(8).toString(),
                    elem.getCell(10).toString()
                        .replace("number", "DECIMAL")
                        .replace("text", "TEXT")
                        .replace("List Of Values", "LIST"),
                    "pattern for values",
                    comment,
                    "pattern for count"
                )

            }
        }
        listOfItemsId.remove("SP-351756")
    }
    for (elem in sheet) {
        if (result[elem.getCell(8).toString()] !== "") {
            //println(elem.getCell(8).toString())
            var comment = ""
            if (elem.getCell(16) !== null) {
                comment = elem.getCell(16).toString()
            }
            if (categoryName == elem.getCell(3).toString()) {
                if (shortName == "") {
                    shortName = (elem.getCell(3).toString()
                            + "<>" + elem.getCell(4).toString()
                            + "<>" + "Short Name"
                            + "<>" + "<>"
                            + "<>" + "REQUIRED"
                            + "<>" + "TEXT"
                            + "<>" + "CNET_Specific_Short_Name"
                            + "<>" + "1"
                            + "<>" + "0"
                            )
                    catKey = "$categoryName Short Name"
                }
                result[categoryName + elem.getCell(8).toString()] = (elem.getCell(3).toString()
                        + "<>" + elem.getCell(4).toString()
                        + "<>" + elem.getCell(8).toString()
                        + "<>" + "Add(R(\"cnet_common_${elem.getCell(9)}\"));"
                        + "<>" + "ready"
                        + "<>" + elem.getCell(13).toString()
                    .replace("N", "OPTIONAL")
                    .replace("Y", "REQUIRED")
                        + "<>" + elem.getCell(10).toString()
                    .replace("number", "DECIMAL")
                    .replace("text", "TEXT")
                    .replace("List Of Values", "LIST")
                        + "<>" + elem.getCell(9).toString()
                        + "<>" + "1"
                        + "<>" + "0"
                        + "<>" + "pattern for values"
                        + "<>" + comment
                        + "<>" + "pattern for count")


            }
        }

    }
    result[catKey] = shortName
    println(categoryName)
    generateCommonItems(listOfItemsId, listOfCommonIds, result, listOfValuesForCommon)

}

fun getListOfValuesForCategory(row: Sheet, listOfValues: MutableMap<String, String>, categoryName: String) {
    row.forEach {
        if (it.getCell(3).toString() == categoryName) {
            if (it.getCell(14).toString() !== "") {
                if (listOfValues[categoryName + it.getCell(8).toString()] == null) {
                    listOfValues[categoryName + it.getCell(8).toString()] = it.getCell(14).toString()
                } else {
                    listOfValues[categoryName + it.getCell(8).toString()] =
                        listOfValues[categoryName + it.getCell(8).toString()] + "|" + it.getCell(14).toString()
                }
            }
        }
    }
}

fun writeToFile(rowsToWrite: MutableMap<String, String>) {
    val myWorkBook = XSSFWorkbook()
    val myWorkList = myWorkBook.createSheet("seed_request")
    var row = 1
    var column = 0
    val firstRow = arrayOf("{L=0}", "{I=0}", "NAME", "EXPRESSION", "EXPRESSION_STATUS", "#8", "#9", "#12", "#45", "#10", "#13", "#14", "#48")
    myWorkList.createRow(0)

    firstRow.forEach {
        myWorkList.getRow(0).createCell(column).setCellValue(it)
        myWorkList.getRow(0).getCell(column)
        column++
    }
    myWorkList.setAutoFilter(CellRangeAddress(0, 0, 0, 12))
    column = 0

    rowsToWrite.forEach { t, u ->
        myWorkList.createRow(row)
        for (item in u.split("<>")) {
            myWorkList.getRow(row).createCell(column).setCellValue(item.replace("pattern for values", "").replace("pattern for count", ""))
            column++
        }
        row++
        column = 0
    }
    val output = FileOutputStream("./seed_test.xlsx")
    myWorkBook.write(output)
}




