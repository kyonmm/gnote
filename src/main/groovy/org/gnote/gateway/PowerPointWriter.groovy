package org.gnote.gateway

import org.apache.poi.sl.usermodel.AutoNumberingScheme
import org.apache.poi.xslf.usermodel.XMLSlideShow

import java.awt.Color
import java.awt.Rectangle

/**
 * Created by kyon_mm on 2018/01/17.
 */
class PowerPointWriter {

    AsciidocReader reader
    def PowerPointWriter(AsciidocReader asciidocReader) {
        reader = asciidocReader
    }

    void generate(String filePath, String outputPath) {
        def document = reader.read(filePath)
        def ppt = new XMLSlideShow();

        def master =  ppt.getSlideMasters().get(0)

        def topSlide = ppt.createSlide();
        def shape = topSlide.createTextBox();

        def p = shape.addNewTextParagraph();

        def r1 = p.addNewTextRun();
        r1.setText(document.doctitle());
        r1.setFontColor(Color.blue);
        r1.setFontSize(24);
        shape.setAnchor(new Rectangle(200, 200, 600, 600));

        document.blocks.eachWithIndex{chapter, index ->
            // ==
            def slide = ppt.createSlide()
            def header = slide.createTextBox()
            header.setText(chapter.title)
            header.setAnchor(new Rectangle(20, 20, 600, 30));
            def level = chapter.level
            // content or ===
            def content = slide.createTextBox()
            chapter.blocks.eachWithIndex{i, idx ->
                if(i.level == level){
                    if(i.blocks.size() == 0){
                        content.setText(i.content.toString())
                    }
                    else{
                        i.blocks.each{listItem ->
                            def paragraph = content.addNewTextParagraph();
                            paragraph.setBulletAutoNumber(AutoNumberingScheme.alphaLcPeriod, 1)
                            def t = paragraph.addNewTextRun()
                            t.setText(listItem.text)
                        }
                    }
                }
                else{
                    println "1だんまで"
                }
            }
            content.setAnchor(new Rectangle(20, 100, 600, 600));
        }


        FileOutputStream out = new FileOutputStream(outputPath);
        ppt.write(out);
        ppt.close()
        out.close();
        sleep(100)
    }


    void generate(String inputPath){
        //new File("generated.pptx").deleteOnExit()
        generate(inputPath, "generated.pptx")
    }
}
