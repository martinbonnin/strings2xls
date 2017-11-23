package net.mbonnin

import cz.tomaskypta.tools.langtool.exporting.ExportConfig
import cz.tomaskypta.tools.langtool.exporting.ToolExport
import org.junit.Test
import java.io.File

class TestDetector {
    @Test
    fun test() {
        val config = ExportConfig()
        config.inputExportProject = File("./test").absolutePath
        config.outputFile = File("exported.xls").absolutePath
        ToolExport.run(config)
    }
}