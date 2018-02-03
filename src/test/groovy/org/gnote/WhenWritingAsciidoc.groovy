package org.gnote

import org.apache.poi.xslf.usermodel.SlideLayout
import org.apache.poi.xslf.usermodel.XMLSlideShow
import org.apache.poi.xslf.usermodel.XSLFSlide
import org.apache.poi.xslf.usermodel.XSLFSlideLayout
import org.apache.poi.xslf.usermodel.XSLFSlideMaster
import org.apache.poi.xslf.usermodel.XSLFTextBox
import org.apache.poi.xslf.usermodel.XSLFTextParagraph
import org.apache.poi.xslf.usermodel.XSLFTextRun
import org.apache.poi.xslf.usermodel.XSLFTextShape
import org.asciidoctor.Asciidoctor
import org.gnote.gateway.AsciidocReader
import spock.lang.Specification

import java.awt.Color
import java.awt.Rectangle

/**
 * Created by kyon_mm on 2018/01/17.
 */
class WhenWritingAsciidoc extends Specification {

  def "write simple asciidoc"(){
    when:
    def asciidoctor = Asciidoctor.Factory.create()
    def hs =  asciidoctor.readDocumentHeader(new File("example.adoc"))
    then:
    hs.documentTitle.main == "Title"
  }

  def "convert pptx"(){
    given:
    def asciidoctor = Asciidoctor.Factory.create()
    def hs =  asciidoctor.readDocumentHeader(new File("example.adoc"))

    when:
    //create a new empty slide show
    XMLSlideShow ppt = new XMLSlideShow();
    XSLFSlide slide = ppt.createSlide();


    XSLFTextBox shape = slide.createTextBox();
    shape.setText(hs.documentTitle.main)
    shape.setAnchor(new Rectangle(100, 100, 200, 200));
    XSLFTextParagraph p = shape.addNewTextParagraph();

    FileOutputStream out = new FileOutputStream("merged.pptx");
    ppt.write(out);
    ppt.close()
    out.close();

    then:
    true

  }

  def "Main Test"(){
    expect:
    Main.main(inputPath)

    where:
    inputPath << ["example.adoc", "~/introduction-asciidoc.adoc"]
  }


}
