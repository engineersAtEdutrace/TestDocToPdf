/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pptxtopdf;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import javax.imageio.ImageIO;
import org.apache.pdfbox.exceptions.COSVisitorException;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.edit.PDPageContentStream;
import org.apache.pdfbox.pdmodel.graphics.xobject.PDJpeg;
import org.apache.pdfbox.pdmodel.graphics.xobject.PDXObjectImage;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.xslf.usermodel.XMLSlideShow;

import org.apache.poi.xslf.usermodel.XSLFShadow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

/**
 *
 * @author sagar
 */
public class TestPptxToPdf {
    
    public static void main(String[] args) throws FileNotFoundException, IOException, COSVisitorException, OpenXML4JException {
        // TODO code application logic here
        String filepath = "/home/sagar/Desktop/Shareback/test.pptx";
        FileInputStream is = new FileInputStream(filepath);
        XMLSlideShow pptx = new XMLSlideShow(is);
        Dimension pgsize = pptx.getPageSize();

        int idx = 1;
        for (XSLFSlide slide : pptx.getSlides()) {

            BufferedImage img = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB);
            Graphics2D graphics = img.createGraphics();
            // clear the drawing area
            graphics.setPaint(Color.white);
            graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));

            // render
            slide.draw(graphics);

            // save the output
            FileOutputStream out = new FileOutputStream("/home/sagar/Desktop/Shareback/img/slide-" + idx + ".jpg");
            javax.imageio.ImageIO.write(img, "jpg", out);
            out.close();

            idx++;
        }
    
        String someimg = "/home/sagar/Desktop/Shareback/pptx/img/";
        
        PDDocument document  = new PDDocument();
        File file = new File(someimg);
        if(!file.exists())
            file.mkdir();
        
        if(file.isDirectory()){
            for(File f: file.listFiles()){

                InputStream in = new FileInputStream(f);

                BufferedImage bimg = ImageIO.read(in);
                float width = bimg.getWidth();
                float height = bimg.getHeight();
                PDPage page = new PDPage(new PDRectangle(width + 10, height + 10));
                document.addPage(page);
                PDXObjectImage img = new PDJpeg(document, new FileInputStream(f));
                PDPageContentStream contentStream = new PDPageContentStream(document, page);
                contentStream.drawImage(img, 0, 0);
                contentStream.close();
                in.close();
            }

            document.save("/home/sagar/Desktop/Shareback/test-generated-pptx.pdf");
            document.close();
        }
        else{
            System.out.println(someimg+"is not a Directory");
        }
        
    }
    
}
