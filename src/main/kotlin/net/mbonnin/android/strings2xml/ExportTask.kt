package net.mbonnin.android.strings2xml

import com.beust.jcommander.Parameter
import cz.tomaskypta.tools.langtool.exporting.ExportConfig
import cz.tomaskypta.tools.langtool.exporting.ToolExport
import org.gradle.api.DefaultTask
import org.gradle.api.tasks.InputDirectory
import org.gradle.api.tasks.OutputFile
import org.gradle.api.tasks.TaskAction
import java.io.File

class ExportTask(): DefaultTask() {
    @InputDirectory
    lateinit var inputDir: File

    @OutputFile
    lateinit var outputFile: File

    @TaskAction
    fun run() {
        val config = ExportConfig()
        config.inputExportProject = inputDir.absolutePath
        config.outputFile = outputFile.absolutePath
        ToolExport.run(config)
    }
}