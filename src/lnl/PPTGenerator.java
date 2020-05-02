/*
 * $Id$
 * Created on Oct 30, 2012 by beardj
 */
package lnl;

import org.docx4j.TraversalUtil;
import org.docx4j.TraversalUtil.CallbackImpl;
import org.docx4j.XmlUtils;
import org.docx4j.dml.CTTextParagraph;
import org.docx4j.dml.chart.CTPieChart;
import org.docx4j.dml.chart.CTPieSer;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.OpcPackage;
import org.docx4j.openpackaging.packages.PresentationMLPackage;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.DrawingML.Chart;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.PresentationML.SlidePart;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.EmbeddedPackagePart;
import org.docx4j.utils.BufferUtil;
import org.pptx4j.jaxb.Context;
import org.pptx4j.pml.Pic;
import org.pptx4j.pml.Shape;
import org.xlsx4j.sml.Row;
import org.xlsx4j.sml.STCellType;

import javax.xml.bind.JAXBException;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.Map.Entry;

public class PPTGenerator {
    public static final String FRUIT_1 = "Fruit1";
    public static final String FRUIT_2 = "Fruit2";
    public static final String FRUIT_3 = "Fruit3";
    public static final String USER_NAME = "User_Name";
    public static final String PID = "pid";

    public static void main(String[] args) throws Docx4JException, IOException {

        // Inputs
        String inputfilepath = "test2.pptx";
        String chartPartName = "/ppt/charts/chart1.xml";
        String xlsPartName = "/ppt/embeddings/Microsoft_Excel_Worksheet.xlsx";
        String outputfilepath = "C:/ppt" + "/OUT_EditEmbeddedCharts-" + System.currentTimeMillis() + ".pptx";
        String pic_location = "C:/ppt/lnl_logo.png";

        //data source
        String Fruit_1 = "Mango";
        String Fruit_2 = "Orange";
        String Fruit_3 = "Guava";
        String User_Name = "Learn Now Lab";
        String pid = "1234";

        //embedded excel
        Map<String, Integer> topModulesMap = new HashMap<>();
        topModulesMap.put("Mango", 1);
        topModulesMap.put("Orange", 2);
        topModulesMap.put("Guava", 3);

        // Open the PPT template file
        PresentationMLPackage ppt = (PresentationMLPackage) OpcPackage.load(new File(inputfilepath));
        SlidePart slide = (SlidePart) ppt.getParts().get(new PartName("/ppt/slides/slide1.xml"));


        //Add the picture
        File file = new File(pic_location);
        BinaryPartAbstractImage imagePart;
        try {
            imagePart = BinaryPartAbstractImage.createImagePart(ppt, slide, file);
            slide.getJaxbElement().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame().add(createPicture(imagePart.getSourceRelationship().getId()));
        } catch (Exception e1) {
            // TODO Auto-generated catch block
            e1.printStackTrace();
        }

        //All shape classes to edit all texts in the template
        ClassFinder dmlShapeFinder = new ClassFinder(Shape.class);
        new TraversalUtil(slide.getJaxbElement().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame(), dmlShapeFinder);
        if (dmlShapeFinder.results.isEmpty()) return;

        for (int i = 0; i < dmlShapeFinder.results.size(); i++) {
            Shape index_shp = (Shape) dmlShapeFinder.results.get(i);

            //*******************  slot starts***********************

            String buff = XmlUtils.marshaltoString(index_shp.getTxBody().getP().get(0), true, true, Context.jcPML,
                    "http://schemas.openxmlformats.org/presentationml/2006/main", "txBody", CTTextParagraph.class);

            if (buff.indexOf(FRUIT_1) != -1) buff = buff.replaceFirst(FRUIT_1, Fruit_1);
            if (buff.indexOf(FRUIT_2) != -1) buff = buff.replaceFirst(FRUIT_2, Fruit_2);
            if (buff.indexOf(FRUIT_3) != -1) buff = buff.replaceFirst(FRUIT_3, Fruit_3);
            if (buff.indexOf(PID) != -1) buff = buff.replaceFirst(PID, pid);
            if (buff.indexOf(USER_NAME) != -1) buff = buff.replaceFirst(USER_NAME, User_Name);

            CTTextParagraph testtt = new CTTextParagraph();
            try {
                testtt = (CTTextParagraph) XmlUtils.unmarshalString(buff, Context.jcPML, CTTextParagraph.class);
            } catch (JAXBException e) {
                e.printStackTrace();
            }
            index_shp.getTxBody().getP().set(0, testtt);

            // ********************** slot ends*************************

        }

        //******************* Chart  starts***********************

        Chart chart = (Chart) ppt.getParts().get(new PartName(chartPartName));
        List<Object> objects = chart.getJaxbElement().getChart().getPlotArea().getAreaChartOrArea3DChartOrLineChart();

        for (Object object : objects) {

            String ptCount = "<c:ptCount val=\"" + topModulesMap.size() + "\"/>";
            String strRef = "";
            String numRef = "";
            Iterator<String> iterator = topModulesMap.keySet().iterator();
            int counter = 0;
            while (iterator.hasNext()) {
                String key = iterator.next().toString();
                Integer value = topModulesMap.get(key);
                strRef = strRef + "<c:pt idx=\"" + counter + "\">" + "<c:v>" + key + "</c:v>" + "</c:pt>";
                numRef = numRef + "<c:pt idx=\"" + counter + "\">" + "<c:v>" + value + "</c:v>" + "</c:pt>";
                counter += 1;
            }


            String addval = "<p:txBody xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:ns6=\"http://schemas.openxmlformats.org/drawingml/2006/chartDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:ns12=\"http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:ns11=\"http://schemas.openxmlformats.org/drawingml/2006/compatibility\" xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
                    + " <c:idx val=\"0\"/>"
                    + " <c:order val=\"0\"/> "
                    + "<c:tx> "
                    + "<c:strRef> <c:f>Sheet1!$B$1</c:f> <c:strCache> <c:ptCount val=\"1\"/> <c:pt idx=\"0\"> <c:v>EBS Purchasing History</c:v> </c:pt> </c:strCache> </c:strRef> "
                    + "</c:tx> "
                    + "<c:cat> "
                    + "<c:strRef> <c:f>Sheet1!$A$2:$A$" + (topModulesMap.size() + 1) + "</c:f> <c:strCache>" + ptCount + strRef + "</c:strCache> </c:strRef>"
                    + "</c:cat> "
                    + "<c:val>"
                    + " <c:numRef> <c:f>Sheet1!$B$2:$B$" + (topModulesMap.size() + 1) + "</c:f> <c:numCache> <c:formatCode>General</c:formatCode>" + ptCount + numRef + "</c:numCache></c:numRef>"
                    + "</c:val> "
                    + "<c:extLst> <c:ext uri=\"{C3380CC4-5D6E-409C-BE32-E72D297353CC}\"> <c16:uniqueId val=\"{00000000-310B-4ADF-8D12-0B51F5325448}\" xmlns:c16r2=\"http://schemas.microsoft.com/office/drawing/2015/06/chart\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\"/> </c:ext> </c:extLst> "
                    + "</p:txBody>";


            CTPieSer testtt123 = new CTPieSer();

            try {
                testtt123 = (CTPieSer) XmlUtils.unmarshalString(addval, Context.jcPML, CTPieSer.class);
            } catch (JAXBException e) {
                e.printStackTrace();
            }

            ((CTPieChart) object).getSer().set(0, testtt123);
        }

        /*
         * Get the spreadsheet and find the cell values that need to be updated
         */

        EmbeddedPackagePart epp = (EmbeddedPackagePart) ppt.getParts().get(new PartName(xlsPartName));

        if (epp == null) {
            throw new Docx4JException("Could find EmbeddedPackagePart: " + xlsPartName);
        }

        InputStream is = BufferUtil.newInputStream(epp.getBuffer());
        SpreadsheetMLPackage spreadSheet = (SpreadsheetMLPackage) SpreadsheetMLPackage.load(is);
        Map<PartName, Part> partsMap = spreadSheet.getParts().getParts();
        Iterator<Entry<PartName, Part>> it = partsMap.entrySet().iterator();

        while (it.hasNext()) {
            Entry<PartName, Part> pairs = it.next();

            if (partsMap.get(pairs.getKey()) instanceof WorksheetPart) {

                WorksheetPart wsp = (WorksheetPart) partsMap.get(pairs.getKey());
                List<Row> rows = wsp.getJaxbElement().getSheetData().getRow();
                Iterator<String> iterator = topModulesMap.keySet().iterator();

                int counter = 1;
                while (iterator.hasNext()) {
                    String key = iterator.next().toString();
                    Integer value = topModulesMap.get(key);
                    rows.get(counter).getC().get(0).setT(STCellType.STR);
                    rows.get(counter).getC().get(0).setV(key);
                    rows.get(counter).getC().get(1).setT(STCellType.N);
                    rows.get(counter).getC().get(1).setV(value.toString());

                    counter += 1;
                }
            }
        }

        //******************* Chart  ends***********************

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        SaveToZipFile saver = new SaveToZipFile(spreadSheet);
        saver.save(baos);
        epp.setBinaryData(baos.toByteArray());
        ppt.save(new File(outputfilepath));

        System.out.println("\n\n done .. saved " + outputfilepath);
    }


    private static Object createPicture(String relId) throws JAXBException {

        // Create p:pic
        HashMap<String, String> mappings = new HashMap<String, String>();

        mappings.put("id1", "4");
        mappings.put("name", "Picture 3");
        mappings.put("descr", "greentick.png");
        mappings.put("rEmbedId", relId);
        mappings.put("offx", Long.toString(10776230));
        mappings.put("offy", Long.toString(130263));
        mappings.put("extcx", Long.toString(1288648));
        mappings.put("extcy", Long.toString(1139648));

        return XmlUtils.unmarshallFromTemplate(SAMPLE_PICTURE,
                mappings, Context.jcPML, Pic.class);

    }

    private static String SAMPLE_PICTURE =
            "<p:pic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"> "
                    + "<p:nvPicPr>"
                    + "<p:cNvPr id=\"${id1}\" name=\"${name}\" descr=\"${descr}\"/>"
                    + "<p:cNvPicPr>"
                    + "<a:picLocks noChangeAspect=\"1\"/>"
                    + "</p:cNvPicPr>"
                    + "<p:nvPr/>"
                    + "</p:nvPicPr>"
                    + "<p:blipFill>"
                    + "<a:blip r:embed=\"${rEmbedId}\" cstate=\"print\"/>"
                    + "<a:stretch>"
                    + "<a:fillRect/>"
                    + "</a:stretch>"
                    + "</p:blipFill>"
                    + "<p:spPr>"
                    + "<a:xfrm>"
                    + "<a:off x=\"${offx}\" y=\"${offy}\"/>"
                    + "<a:ext cx=\"${extcx}\" cy=\"${extcy}\"/>"
                    + "</a:xfrm>"
                    + "<a:prstGeom prst=\"rect\">"
                    + "<a:avLst/>"
                    + "</a:prstGeom>"
                    + "</p:spPr>"
                    + "</p:pic>";


    static class ClassFinder extends CallbackImpl {

        protected Class<?> typeToFind;

        public ClassFinder(Class<?> typeToFind) {
            this.typeToFind = typeToFind;
        }

        public List<Object> results = new ArrayList<Object>();

        @Override
        public List<Object> apply(Object o) {

            // Adapt as required
            if (o.getClass().equals(typeToFind)) {
                results.add(o);
            }
            return null;
        }

        public List<Object> getChildren(Object o) {

            if (o instanceof org.pptx4j.pml.CTGraphicalObjectFrame) {
                org.docx4j.dml.Graphic graphic = ((org.pptx4j.pml.CTGraphicalObjectFrame) o).getGraphic();
                if (graphic != null
                        && graphic.getGraphicData() != null) {
                    return graphic.getGraphicData().getAny();
                } else {
                    return null;
                }
            }

            return TraversalUtil.getChildrenImpl(o);
        }

    }

}