package com.seal.tool;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObject;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.List;
import java.util.Random;

/**
 * @Author alan.wang
 * @date: 2022-10-10 08:58
 */
public class WordSealUtil {

    public static void main(String[] args) throws Exception {


        String inputFile = args[0]; //D:\soft\doc\(综合集控)集控操作工操作规程.docx
        String targetFile = args[1];//D:\soft\doc\SealInWord.docx
        String sealImage = args[2];//D:\soft\doc\controlled.png

        sealInWord(inputFile,targetFile,sealImage,"(签字/盖章)",0,0,400,-50,true);
    }
    /**
     * <b> Word中添加图章
     * </b><br><br><i>Description</i> :
     * String srcPath, 源Word路径
     * String storePath, 添加图章后的路径
     * String sealPath, 图章路径（即图片）
     * tString abText, 在Word中盖图章的标识字符串，如：(签字/盖章)
     * int width, 图章宽度
     * int height, 图章高度
     * int leftOffset, 图章在编辑段落向左偏移量
     * int topOffset, 图章在编辑段落向上偏移量
     * boolean behind，图章是否在文字下面
     * @return void
     * <br><br>Date: 2019/12/26 15:12     <br>Author : dxl
     */
    public static void sealInWord(String srcPath, String storePath,String sealPath,String tabText,
                                  int width, int height, int leftOffset, int topOffset, boolean behind) throws Exception {
        File fileTem = new File(srcPath);
        InputStream is = new FileInputStream(fileTem);
        XWPFDocument document = new XWPFDocument(is);

        List<XWPFParagraph> paragraphs = document.getParagraphs();
        XWPFRun targetRun = null;
        for(XWPFParagraph  paragraph : paragraphs){
            if(!"".equals(paragraph.getText()) ){
                List<XWPFRun> runs = paragraph.getRuns();
                targetRun = runs.get(runs.size()-1);
                break;
            }
        }
//        List<XWPFHeader> headers = document.getHeaderList();
//        for(XWPFHeader  header : headers){
//            List<XWPFParagraph> hparagraphs = header.getParagraphs();
//            for(XWPFParagraph  paragraph : hparagraphs){
//                if(!"".equals(paragraph.getText()) ){
//                    List<XWPFRun> runs = paragraph.getRuns();
//                    targetRun = runs.get(runs.size()-1);
//                    break;
//                }
//            }
//        }
        if(targetRun != null){
            InputStream in = new FileInputStream(sealPath);//设置图片路径
            if(width <= 0){
                width = 100;
            }
            if(height <= 0){
                height = 100;
            }
            //创建Random类对象
            Random random = new Random();
            //产生随机数
            int number = random.nextInt(999) + 1;
            targetRun.addPicture(in, Document.PICTURE_TYPE_PNG, "Seal"+number, Units.toEMU(width), Units.toEMU(height));
            in.close();
            // 2. 获取到图片数据
            CTDrawing drawing = targetRun.getCTR().getDrawingArray(0);
            CTGraphicalObject graphicalobject = drawing.getInlineArray(0).getGraphic();

            //拿到新插入的图片替换添加CTAnchor 设置浮动属性 删除inline属性
            CTAnchor anchor = getAnchorWithGraphic(graphicalobject, "Seal"+number,
                    Units.toEMU(width), Units.toEMU(height),//图片大小
                    Units.toEMU(leftOffset), Units.toEMU(topOffset), behind);//相对当前段落位置 需要计算段落已有内容的左偏移
            drawing.setAnchorArray(new CTAnchor[]{anchor});//添加浮动属性
            drawing.removeInline(0);//删除行内属性
        }
        document.write(new FileOutputStream(storePath));
        document.close();
    }
    /**
     * @param ctGraphicalObject 图片数据
     * @param deskFileName      图片描述
     * @param width             宽
     * @param height            高
     * @param leftOffset        水平偏移 left
     * @param topOffset         垂直偏移 top
     * @param behind            文字上方，文字下方
     * @return
     * @throws Exception
     */
    public static CTAnchor getAnchorWithGraphic(CTGraphicalObject ctGraphicalObject,
                                                String deskFileName, int width, int height,
                                                int leftOffset, int topOffset, boolean behind) {
        System.out.println(">>width>>"+width+"; >>height>>>>"+height);
        String anchorXML =
                "<wp:anchor xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
                        + "simplePos=\"0\" relativeHeight=\"0\" behindDoc=\"" + ((behind) ? 1 : 0) + "\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">"
                        + "<wp:simplePos x=\"100\" y=\"200\"/>"
                        + "<wp:positionH relativeFrom=\"column\">"
                        + "<wp:posOffset>" + leftOffset + "</wp:posOffset>"
                        + "</wp:positionH>"
                        + "<wp:positionV relativeFrom=\"paragraph\">"
                        + "<wp:posOffset>" + topOffset + "</wp:posOffset>" +
                        "</wp:positionV>"
                        + "<wp:extent cx=\"" + width + "\" cy=\"" + height + "\"/>"
                        + "<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>"
                        + "<wp:wrapNone/>"
                        + "<wp:docPr id=\"1\" name=\"Drawing 0\" descr=\"" + deskFileName + "\"/><wp:cNvGraphicFramePr/>"
                        + "</wp:anchor>";

        CTDrawing drawing = null;
        try {
            drawing = CTDrawing.Factory.parse(anchorXML);
        } catch (XmlException e) {
            e.printStackTrace();
        }
        CTAnchor anchor = drawing.getAnchorArray(0);
        anchor.setGraphic(ctGraphicalObject);
        return anchor;
    }

}
