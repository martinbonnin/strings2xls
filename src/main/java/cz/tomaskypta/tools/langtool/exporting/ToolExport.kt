package cz.tomaskypta.tools.langtool.exporting

import java.io.*
import java.util.*
import javax.xml.parsers.DocumentBuilder
import javax.xml.parsers.DocumentBuilderFactory
import javax.xml.parsers.ParserConfigurationException

import org.apache.commons.lang3.StringUtils
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.util.CellRangeAddress
import org.w3c.dom.Document
import org.w3c.dom.Node
import org.w3c.dom.NodeList
import org.xml.sax.SAXException


class ToolExport @Throws(ParserConfigurationException::class)
constructor(out: PrintStream?) {

    private val builder: DocumentBuilder
    private var outExcelFile: File? = null
    private var project: String? = null
    private var keysIndex: Map<String, Int>? = null
    private val out: PrintStream
    private var mConfig: ExportConfig? = null
    private val sAllowedFiles = HashSet<String>()

    init {
        sAllowedFiles.add("strings.xml")
    }

    init {
        val dbf = DocumentBuilderFactory.newInstance()
        builder = dbf.newDocumentBuilder()
        this.out = out ?: System.out
    }

    @Throws(SAXException::class, IOException::class)
    private fun export(project: File) {
        val res = findResourceDir(project)
        if (res == null) {
            System.err.println("Cannot find resource directory.")
            return
        }


        // make sure we have DIR_VALUES first so we have the keys
        val subdirList = res.listFiles()
        subdirList.sortBy {it.name}

        for (dir in subdirList) {
            if (!dir.isDirectory || !dir.name.startsWith(DIR_VALUES)) {
                continue
            }
            val dirName = dir.name
            if (dirName == DIR_VALUES) {
                keysIndex = exportDefLang(dir)
            } else {
                val index = dirName.indexOf('-')
                if (index == -1)
                    continue
                val lang = dirName.substring(index + 1)
                exportLang(lang, dir)
            }
        }
    }

    private fun findResourceDir(project: File): File? {
        val availableResDirs = LinkedList<File>()
        for (potentialResDir in POTENTIAL_RES_DIRS) {
            val res = File(project, potentialResDir)
            if (res.exists()) {
                availableResDirs.add(res)
            }
        }
        return if (!availableResDirs.isEmpty()) {
            availableResDirs[0]
        } else null
    }

    @Throws(IOException::class, SAXException::class)
    private fun exportLang(lang: String, valueDir: File) {
        for (fileName in sAllowedFiles) {
            val stringFile = File(valueDir, fileName)
            if (!stringFile.exists()) {
                continue
            }
            exportLangToExcel(project, lang, stringFile, getStrings(stringFile), outExcelFile, keysIndex)
        }
    }

    @Throws(IOException::class, SAXException::class)
    private fun exportDefLang(valueDir: File): Map<String, Int> {
        val keys = HashMap<String, Int>()
        val wb = HSSFWorkbook()

        val sheet: HSSFSheet
        sheet = wb.createSheet(project!!)
        var rowIndex = 0
        sheet.createRow(rowIndex++)
        createTilte(wb, sheet)
        addLang2Tilte(wb, sheet, "default")
        sheet.createFreezePane(1, 1)

        val outFile = FileOutputStream(outExcelFile!!)
        wb.write(outFile)
        outFile.close()

        for (fileName in sAllowedFiles) {
            val stringFile = File(valueDir, fileName)
            if (!stringFile.exists()) {
                continue
            }
            keys.putAll(exportDefLangToExcel(rowIndex, project!!, stringFile, getStrings(stringFile), outExcelFile!!))
        }


        return keys
    }

    @Throws(SAXException::class, IOException::class)
    private fun getStrings(f: File): NodeList {
        val dom = builder.parse(f)
        return dom.documentElement.childNodes
    }


    @Throws(FileNotFoundException::class, IOException::class)
    private fun exportDefLangToExcel(rowIndex: Int, project: String, src: File, strings: NodeList, f: File): Map<String, Int> {
        var rowIndex = rowIndex
        out.println()
        out.println("Start processing DEFAULT language " + src.name)

        val keys = HashMap<String, Int>()

        val wb = HSSFWorkbook(FileInputStream(f))

        val commentStyle = createCommentStyle(wb)
        val plurarStyle = createPlurarStyle(wb)
        val keyStyle = createKeyStyle(wb)
        val textStyle = createTextStyle(wb)

        val sheet = wb.getSheet(project)


        for (i in 0 until strings.length) {
            val item = strings.item(i)
            if (item.nodeType == Node.TEXT_NODE) {

            }
            if (item.nodeType == Node.COMMENT_NODE) {
                val row = sheet.createRow(rowIndex++)
                val cell = row.createCell(0)
                cell.setCellValue(String.format("/** %s **/", item.textContent))
                cell.setCellStyle(commentStyle)

                sheet.addMergedRegion(CellRangeAddress(row.rowNum, row.rowNum, 0, 255))
            }

            if ("string" == item.nodeName) {
                val translatable = item.attributes.getNamedItem("translatable")
                if (translatable != null && "false" == translatable.nodeValue) {
                    continue
                }
                val key = item.attributes.getNamedItem("name").nodeValue
                if (mConfig!!.isIgnoredKey(key)!!) {
                    continue
                }
                keys.put(key, rowIndex)

                val row = sheet.createRow(rowIndex++)

                var cell = row.createCell(0)
                cell.setCellValue(key)
                cell.setCellStyle(keyStyle)

                cell = row.createCell(1)
                cell.setCellStyle(textStyle)
                cell.setCellValue(item.textContent)
            } else if ("plurals" == item.nodeName) {
                val key = item.attributes.getNamedItem("name").nodeValue
                if (mConfig!!.isIgnoredKey(key)!!) {
                    continue
                }

                val row = sheet.createRow(rowIndex++)
                val cell = row.createCell(0)
                cell.setCellValue(String.format("//plurals: %s", key))
                cell.setCellStyle(plurarStyle)

                val items = item.childNodes
                for (j in 0 until items.length) {
                    val plurarItem = items.item(j)
                    if ("item" == plurarItem.nodeName) {
                        val itemKey = key + "#" + plurarItem.attributes.getNamedItem("quantity").nodeValue
                        keys.put(itemKey, rowIndex)

                        val itemRow = sheet.createRow(rowIndex++)

                        var itemCell = itemRow.createCell(0)
                        itemCell.setCellValue(itemKey)
                        itemCell.setCellStyle(keyStyle)

                        itemCell = itemRow.createCell(1)
                        itemCell.setCellStyle(textStyle)
                        itemCell.setCellValue(plurarItem.textContent)
                    }
                }
            } else if ("string-array" == item.nodeName) {
                val key = item.attributes.getNamedItem("name").nodeValue
                if (mConfig!!.isIgnoredKey(key)!!) {
                    continue
                }
                val arrayItems = item.childNodes
                var j = 0
                var k = 0
                while (j < arrayItems.length) {
                    val arrayItem = arrayItems.item(j)
                    if ("item" == arrayItem.nodeName) {
                        val itemKey = key + "[" + k++ + "]"
                        keys.put(itemKey, rowIndex)

                        val itemRow = sheet.createRow(rowIndex++)

                        var itemCell = itemRow.createCell(0)
                        itemCell.setCellValue(itemKey)
                        itemCell.setCellStyle(keyStyle)

                        itemCell = itemRow.createCell(1)
                        itemCell.setCellStyle(textStyle)
                        itemCell.setCellValue(arrayItem.textContent)
                    }
                    j++
                }
            }
        }

        val outFile = FileOutputStream(f)
        wb.write(outFile)
        outFile.close()

        out.println("DEFAULT language was precessed")
        return keys
    }

    @Throws(FileNotFoundException::class, IOException::class)
    private fun exportLangToExcel(project: String?, lang: String, src: File, strings: NodeList, f: File?, keysIndex: Map<String, Int>?) {
        out.println()
        out.println(String.format("Start processing: '%s'", lang) + " " + src.name)
        val missedKeys = HashSet(keysIndex!!.keys)

        val wb = HSSFWorkbook(FileInputStream(f!!))

        val textStyle = createTextStyle(wb)

        val sheet = wb.getSheet(project)
        addLang2Tilte(wb, sheet, lang)

        val titleRow = sheet.getRow(0)
        val lastColumnIdx = titleRow.lastCellNum.toInt() - 1

        for (i in 0 until strings.length) {
            val item = strings.item(i)

            if ("string" == item.nodeName) {
                val translatable = item.attributes.getNamedItem("translatable")
                if (translatable != null && "false" == translatable.nodeValue) {
                    continue
                }
                val key = item.attributes.getNamedItem("name").nodeValue
                val index = keysIndex[key]
                if (index == null) {
                    out.println("\t" + key + " - row does not exist")
                    continue
                }

                missedKeys.remove(key)
                val row = sheet.getRow(index)

                val cell = row.createCell(lastColumnIdx)
                cell.setCellValue(item.textContent)
                cell.setCellStyle(textStyle)
            } else if ("plurals" == item.nodeName) {
                var key = item.attributes.getNamedItem("name").nodeValue
                val plurarName = key

                val items = item.childNodes
                for (j in 0 until items.length) {
                    val plurarItem = items.item(j)
                    if ("item" == plurarItem.nodeName) {
                        key = plurarName + "#" + plurarItem.attributes.getNamedItem("quantity").nodeValue
                        val index = keysIndex[key]
                        if (index == null) {
                            out.println("\t" + key + " - row does not exist")
                            continue
                        }
                        missedKeys.remove(key)

                        val row = sheet.getRow(index)

                        val cell = row.createCell(lastColumnIdx)
                        cell.setCellValue(plurarItem.textContent)
                        cell.setCellStyle(textStyle)
                    }
                }
            } else if ("string-array" == item.nodeName) {
                val key = item.attributes.getNamedItem("name").nodeValue
                val arrayItems = item.childNodes
                var j = 0
                var k = 0
                while (j < arrayItems.length) {
                    val arrayItem = arrayItems.item(j)
                    if ("item" == arrayItem.nodeName) {
                        val itemKey = key + "[" + k++ + "]"
                        val rowIndex = keysIndex[itemKey]
                        if (rowIndex == null) {
                            out.println("\t" + key + " - row does not exist")
                            j++
                            continue
                        }
                        missedKeys.remove(key)

                        val itemRow = sheet.getRow(rowIndex)

                        val cell = itemRow.createCell(lastColumnIdx)
                        cell.setCellValue(arrayItem.textContent)
                        cell.setCellStyle(textStyle)
                    }
                    j++
                }
            }
        }

        val missedStyle = createMissedStyle(wb)

        if (!missedKeys.isEmpty()) {
            out.println("  MISSED KEYS:")
        }
        for (missedKey in missedKeys) {
            out.println("\t" + missedKey)
            val index = keysIndex[missedKey]
            val row = sheet.getRow(index!!)
            val cell = row.createCell(row.lastCellNum.toInt())
            cell.setCellStyle(missedStyle)
        }

        val outStream = FileOutputStream(f)
        wb.write(outStream)
        outStream.close()

        if (missedKeys.isEmpty()) {
            out.println(String.format("'%s' was processed", lang))
        } else {
            out.println(String.format("'%s' was processed with MISSED KEYS - %d", lang, missedKeys.size))
        }
    }

    companion object {

        private val DIR_VALUES = "values"
        private val POTENTIAL_RES_DIRS = arrayOf("res", "src/main/res")

        @Throws(SAXException::class, IOException::class, ParserConfigurationException::class)
        fun run(config: ExportConfig) {
            run(null, config)
        }

        @Throws(SAXException::class, IOException::class, ParserConfigurationException::class)
        fun run(out: PrintStream?, config: ExportConfig) {
            val tool = ToolExport(out)
            if (StringUtils.isEmpty(config.inputExportProject)) {
                tool.out.println("Cannot export, missing config")
                return
            }
            val project = File(config.inputExportProject)
            if (StringUtils.isEmpty(config.outputFile)) {
                config.outputFile = "exported_strings_" + System.currentTimeMillis() + ".xls"
            }
            tool.outExcelFile = File(config.outputFile)
            tool.project = project.name
            tool.mConfig = config
            tool.sAllowedFiles.addAll(config.additionalResources)
            tool.export(project)
        }

        private fun createTilteStyle(wb: HSSFWorkbook): HSSFCellStyle {
            val bold = wb.createFont()
            bold.boldweight = HSSFFont.BOLDWEIGHT_BOLD

            val style = wb.createCellStyle()
            style.setFont(bold)
            style.fillForegroundColor = HSSFColor.GREY_25_PERCENT.index
            style.fillPattern = HSSFCellStyle.SOLID_FOREGROUND
            style.alignment = HSSFCellStyle.ALIGN_CENTER
            style.wrapText = true

            return style
        }

        private fun createCommentStyle(wb: HSSFWorkbook): HSSFCellStyle {

            val commentFont = wb.createFont()
            commentFont.color = HSSFColor.GREEN.index
            commentFont.italic = true
            commentFont.fontHeightInPoints = 12.toShort()

            val commentStyle = wb.createCellStyle()
            commentStyle.setFont(commentFont)
            return commentStyle
        }

        private fun createPlurarStyle(wb: HSSFWorkbook): HSSFCellStyle {

            val commentFont = wb.createFont()
            commentFont.color = HSSFColor.GREY_50_PERCENT.index
            commentFont.italic = true
            commentFont.fontHeightInPoints = 12.toShort()

            val commentStyle = wb.createCellStyle()
            commentStyle.setFont(commentFont)
            return commentStyle
        }

        private fun createKeyStyle(wb: HSSFWorkbook): HSSFCellStyle {
            val bold = wb.createFont()
            bold.boldweight = HSSFFont.BOLDWEIGHT_BOLD
            bold.fontHeightInPoints = 11.toShort()

            val keyStyle = wb.createCellStyle()
            keyStyle.setFont(bold)

            return keyStyle
        }

        private fun createTextStyle(wb: HSSFWorkbook): HSSFCellStyle {
            val plain = wb.createFont()
            plain.fontHeightInPoints = 12.toShort()

            val textStyle = wb.createCellStyle()
            textStyle.setFont(plain)

            return textStyle
        }

        private fun createMissedStyle(wb: HSSFWorkbook): HSSFCellStyle {

            val style = wb.createCellStyle()
            style.fillForegroundColor = HSSFColor.LIGHT_ORANGE.index
            style.fillPattern = HSSFCellStyle.SOLID_FOREGROUND

            return style
        }

        private fun createTilte(wb: HSSFWorkbook, sheet: HSSFSheet) {
            val titleRow = sheet.getRow(0)

            val cell = titleRow.createCell(0)
            cell.setCellStyle(createTilteStyle(wb))
            cell.setCellValue("KEY")

            sheet.setColumnWidth(cell.columnIndex, 40 * 256)
        }

        private fun addLang2Tilte(wb: HSSFWorkbook, sheet: HSSFSheet, lang: String) {
            val titleRow = sheet.getRow(0)
            val lastCell = titleRow.getCell(titleRow.lastCellNum.toInt() - 1)
            if (lang == lastCell.stringCellValue) {
                // language column already exists
                return
            }
            val cell = titleRow.createCell(titleRow.lastCellNum.toInt())
            cell.setCellStyle(createTilteStyle(wb))
            cell.setCellValue(lang)

            sheet.setColumnWidth(cell.columnIndex, 60 * 256)
        }
    }
}
