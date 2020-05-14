import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class PptTest {
    public static void main(String[] args) throws IOException {
        //创建一个ppt  顶层对象 不管是读取ppt还是写入一个ppt
        XMLSlideShow ppt = new XMLSlideShow();
        //获取当前页
        Dimension pageSize = ppt.getPageSize();

        pageSize.setSize(800,700);


        //获取幻灯片主题列表：
        List<XSLFSlideMaster> slideMasters = ppt.getSlideMasters();
        //获取幻灯片的布局样式
        XSLFSlideLayout layout = slideMasters.get(0).getLayout(SlideLayout.TITLE_AND_CONTENT);
        //通过布局样式创建幻灯片
        XSLFSlide slide = ppt.createSlide(layout);
        // 创建一张无样式的幻灯片
//        XSLFSlide slide = ppt.createSlide();

        //通过当前幻灯片的布局找到第一个空白区：
        XSLFTextShape placeholder = slide.getPlaceholder(0);
        XSLFTextRun title = placeholder.setText("成都智互联科技有限公司");
        XSLFTextShape content = slide.getPlaceholder(1);
        //   投影片中现有的文字
        content.clearText();
        content.setText("图片区");

        // reading an image
        File image = new File("C:\\Users\\Administrator\\Desktop\\111.png");
        //获取图片信息：
        BufferedImage img = ImageIO.read(image);
        // converting it into a byte array
        byte[] picture = IOUtils.toByteArray(new FileInputStream(image));

        // adding the image to the presentation
        XSLFPictureData idx = ppt.addPicture(picture, PictureData.PictureType.PNG);

        // creating a slide with given picture on it
        XSLFPictureShape pic = slide.createPicture(idx);
        //设置当前图片在ppt中的位置，以及图片的宽高
        pic.setAnchor(new java.awt.Rectangle(360, 200, img.getWidth(), img.getHeight()));
        // creating a file object
        File file = new File("C:\\Users\\Administrator\\Desktop\\AddImageToPPT.pptx");
        FileOutputStream out = new FileOutputStream(file);
        // saving the changes to a file
        ppt.write(out);
        System.out.println("image added successfully");
        out.close();


    }



}
