package net.mbonnin.android.strings2xml

import org.gradle.api.Plugin
import org.gradle.api.Project

open class Strings2XlsPlugin : Plugin<Project> {
    override fun apply(project: Project) {
        val xlsFile = project.file("./build/outputs/xls/strings.xls").absoluteFile
        val intputDir = project.file(".")

        val task = project.tasks.create("strings2Xls", ExportTask::class.java) {
            it.inputDir = intputDir
            it.outputFile = xlsFile
        }

        task.inputs.dir(intputDir)
        task.outputs.file(xlsFile)
    }
}