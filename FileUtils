import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.tools.zip.ZipEntry;
import org.apache.tools.zip.ZipOutputStream;

public class FileUtils {

	/**
	 * 将html内容写入word
	 * @param html
	 * @param toFile
	 * @return
	 * @throws Exception
	 */
    public static boolean writeHtml2Word(String html, File toFile) throws Exception {
        boolean flag = false;
        ByteArrayInputStream bais = null;
        FileOutputStream fos = null;
        try {
        	 //去掉html中的注释(<!-- xxxx -->)
      	     html = html.replaceAll("<!\\-\\-.*?\\-\\->", "");
             byte b[] = html.getBytes();
             bais = new ByteArrayInputStream(b);
             POIFSFileSystem poifs = new POIFSFileSystem();
             DirectoryEntry directory = poifs.getRoot();
             directory.createDocument("WordDocument", bais);
             fos = new FileOutputStream(toFile);
             poifs.writeFilesystem(fos);
             bais.close();
             fos.close();
             flag = true;
        } catch (IOException e) {
               e.printStackTrace();
        } finally {
               if(fos != null) fos.close();
               if(bais != null) bais.close();
        }
        return flag;
    }
    
    /**
     * 压缩文件
     * @param fileList
     * @param zipFile
     * @throws IOException
     */
    public static void zipFiles(List<File> fileList, File zipFile) throws IOException{
    	if(zipFile.isDirectory() || !zipFile.getName().endsWith(".zip")) return;
    	
    	ZipOutputStream out = null;
    	FileInputStream fis = null;
		try {
			out = new ZipOutputStream(new FileOutputStream(zipFile));
			//设置压缩方法和压缩级别，级别是0-9，9压缩比例最高
			out.setMethod(ZipOutputStream.DEFLATED);
			out.setLevel(8);
	    	byte[] buffer = new byte[1024];
			for(File file : fileList){
				fis = new FileInputStream(file);   
	            out.putNextEntry(new ZipEntry(file.getName()));     
	            //out.setEncoding("GBK");   
	            int len;      
	            while ((len = fis.read(buffer)) > 0) {   
	                out.write(buffer, 0, len);   
	            }   
	            out.closeEntry();   
	            fis.close(); 
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if(null!=fis) fis.close();
			if(null!=out) out.close();
		}

    }
