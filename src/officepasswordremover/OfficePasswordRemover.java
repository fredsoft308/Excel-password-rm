package officepasswordremover;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class OfficePasswordRemover {

    final String settingsWord = "word/settings.xml";
    final String sheetsExcel = "xl/worksheets/sheet";
    final String workbookExcel = "xl/workbook.xml";

    public OfficePasswordRemover(File file) {
        this.readFile(file);
    }

    private void readFile(File protectedFile) {
        try {

            HashMap<String, File> list = new HashMap<>();

            // https://stackoverflow.com/a/15667326
            ZipFile zipFile = new ZipFile(protectedFile);
            Enumeration<? extends ZipEntry> entries = zipFile.entries();

            while (entries.hasMoreElements()) {
                ZipEntry entry = (ZipEntry) entries.nextElement();

                /**
                 * Excel
                 */
                if (entry.getName().startsWith(sheetsExcel)) {
                    File temp = this.removeNode(zipFile, entry, "sheetProtection");
                    list.put(entry.getName(), temp);
                }
                if (entry.getName().equals(workbookExcel)) {
                    File temp = this.removeNode(zipFile, entry, "workbookProtection");
                    list.put(entry.getName(), temp);
                }
                

                /**
                 * Word
                 */
                if (entry.getName().equals(this.settingsWord)) {
                    File temp = this.removeNode(zipFile, entry, "w:documentProtection");
                    list.put(entry.getName(), temp);
                }

            }

            zipFile.close();
            for (Map.Entry<String, File> entry : list.entrySet()) {
                String key = entry.getKey();
                File tempFile = entry.getValue();

                if (tempFile != null) {
                    this.replaceFile(protectedFile.toPath(), tempFile.toPath(), key);
                    tempFile.delete();
                }
            }

        } catch (Exception ex) {
            Logger.getLogger(OfficePasswordRemover.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private void replaceFile(Path zipFile, Path replacementFile, String initialFile) {
        // https://stackoverflow.com/a/43836969
        // https://stackoverflow.com/a/13787907
        // https://stackoverflow.com/a/14648600
        try (FileSystem fs = FileSystems.newFileSystem(zipFile)) {
            Path fileInsideZipPath = fs.getPath(initialFile);
            Files.copy(replacementFile, fileInsideZipPath, StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    private File removeNode(ZipFile zipFile, ZipEntry entry, String nodeName) throws Exception {
        InputStream stream = zipFile.getInputStream(entry);

        // Parse XML
        DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
        dbFactory.setNamespaceAware(true);
        DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
        Document doc = dBuilder.parse(stream);
        stream.close();

        // Remove Node if available
        // https://examples.javacodegeeks.com/core-java/xml/dom/remove-node-from-dom-document/
        Element element = (Element) doc.getElementsByTagName(nodeName).item(0);
        if (element != null) {
            element.getParentNode().removeChild(element);

            // Create temp File for new settings file
            File tempFile = File.createTempFile("office", "test");
            tempFile.deleteOnExit();
            // Save XML
            // https://stackoverflow.com/a/4561785
            Transformer transformer = TransformerFactory.newInstance().newTransformer();
            Result output = new StreamResult(tempFile);
            Source input = new DOMSource(doc);
            transformer.transform(input, output);

            return tempFile;

        }
        return null;
    }

}
