package org.gnote.gateway

import org.apache.poi.sl.usermodel.AutoNumberingScheme
import org.apache.poi.xslf.usermodel.XMLSlideShow
import org.apache.poi.xslf.usermodel.XSLFSlide

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
        /**
         * Title Slide
         Picture with Caption
         Title and Vertical Text
         Comparison
         Blank
         Vertical Title and Text
         Title and Content
         Title Only
         Section Header
         Two Content
         Content with Caption
         */

        def topSlide = ppt.createSlide(master.getLayout("Title Slide"));
        //topSlide.getPlaceholders().each{it.clearText()}
        def title = topSlide.getPlaceholder(0)
        title.setText(document.doctitle())

        // toc
        def toc = ppt.createSlide(master.getLayout("Title and Content"))
        toc.getPlaceholders().each { it.clearText() }
        def tocHeader = toc.getPlaceholder(0)
        tocHeader.setText("Agenda")
        document.blocks.eachWithIndex { chapter, index ->
            def content = toc.getPlaceholder(1)
            content.appendText(chapter.title, true)
        }



        document.blocks.eachWithIndex{chapter, index ->
            // == section
            def sectionSlide = ppt.createSlide(master.getLayout("Section Header"))
            //sectionSlide.getPlaceholders().each{it.clearText()}
            def sectionHeader = sectionSlide.getPlaceholder(0)
            sectionHeader.setText(chapter.title)
            // ==
            def slide = ppt.createSlide(master.getLayout("Title and Content"))
            slide.getPlaceholders().each{it.clearText()}
            def header = slide.getPlaceholder(0)
            header.setText(chapter.title)
            def level = chapter.level
            // content or ===
            def content = slide.getPlaceholder(1)
            chapter.blocks.eachWithIndex{i, idx ->
                if(i.level == level){
                    if(i.blocks.size() == 0){
                        content.appendText(i.content.toString(), true)
                    }
                    else{
                        i.blocks.each{listItem ->
                            content.appendText(listItem.text, true)
                        }
                    }
                }
                else{
                    println "1だんまで"
                }
            }
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
