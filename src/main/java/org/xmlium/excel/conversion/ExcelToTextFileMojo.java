package org.xmlium.excel.conversion;
import org.apache.maven.plugin.AbstractMojo;
import org.apache.maven.plugin.MojoExecutionException;
import org.apache.maven.plugins.annotations.Mojo;
import org.apache.maven.plugins.annotations.Parameter;

@Mojo(name = "generate", threadSafe = true)
public class ExcelToTextFileMojo extends AbstractMojo {

    @Parameter(defaultValue = "${project.basedir.path}", readonly = true, required = true)
    private String parentBasedir;

    public void setParentBasedir(String parentBasedir) {
        this.parentBasedir = parentBasedir;
    }

    @Override
    public void execute() throws MojoExecutionException {
        getLog().info("***************************");
        getLog().info("*****  Maven Plugin Xslx To Text File Conversion   ******");
        getLog().info("***************************");
        ExcelToTextFile excelToTextFile = new ExcelToTextFile(getLog(), parentBasedir);
        try {
            excelToTextFile.generateTextFilesFromExcelFile();
        } catch (Exception e) {
            throw new MojoExecutionException("Failed to perform files conversion : " + e.getMessage(), e);
        }
        getLog().info("***************************");
    }
}