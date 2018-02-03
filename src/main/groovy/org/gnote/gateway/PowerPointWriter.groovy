package org.gnote.gateway

import org.apache.poi.sl.usermodel.AutoNumberingScheme
import org.apache.poi.sl.usermodel.PictureData
import org.apache.poi.util.IOUtils
import org.apache.poi.xslf.usermodel.XMLSlideShow
import org.apache.poi.xslf.usermodel.XSLFSlideMaster
import org.apache.poi.xslf.usermodel.XSLFTextShape
import org.asciidoctor.ast.AbstractBlock
import org.asciidoctor.ast.BlockImpl
import org.asciidoctor.ast.Document
import org.asciidoctor.ast.ListImpl
import org.asciidoctor.ast.ListItemImpl
import org.jsoup.Jsoup

/**
 * Created by kyon_mm on 2018/01/17.
 */
class PowerPointWriter {

  AsciidocReader reader

  def PowerPointWriter(AsciidocReader asciidocReader) {
    reader = asciidocReader
  }
  def appender = { XSLFTextShape shape, List<AbstractBlock> list, boolean ordered, int indent ->
    list.each {
      if (it.hasProperty("text")) {
        def p = shape.addNewTextParagraph()
        if (ordered) {
          p.setBulletAutoNumber(AutoNumberingScheme.arabicPeriod, 1)
        }
        p.setIndentLevel(indent)
        if(it.text.startsWith("<a href=\"")){
          def xml = Jsoup.parse(it.text)
          def link = xml.select("a")
          def linkPath = link.attr("href")
          def text = link.text()
          def t = p.addNewTextRun()
          t.setText(text)
          def hyperLink = t.createHyperlink()
          hyperLink.linkToUrl(linkPath)
        }
        else{
          def t = p.addNewTextRun()
          t.setText(it.text)
        }
      }
      if (0 < it.getBlocks().size()) {
        def isListImpl = it instanceof ListImpl
        appender(shape, it.getBlocks(), isListImpl ? it.getContext().contains("olist") : ordered, isListImpl ? indent : indent + 1)
      }
    }
  }
  def appenderItem = { XSLFTextShape shape, ListItemImpl list, boolean ordered, int indent ->
    list.each {
      if (it.hasProperty("text")) {
        def p = shape.addNewTextParagraph()
        if (ordered) {
          p.setBulletAutoNumber(AutoNumberingScheme.arabicPeriod, 1)
        }
        p.setIndentLevel(indent)
        def t = p.addNewTextRun()
        t.setText(it.text)
      }
      if (0 < it.getBlocks().size()) {
        def isListImpl = it instanceof ListImpl
        appender(shape, it.getBlocks(), isListImpl ? it.getContext().contains("olist") : ordered, isListImpl ? indent : indent + 1)
      }
    }
  }

  void generate(String filePath, String outputPath) {
    def f = new File(filePath.replace(/~/, System.getProperty("user.home")))

    def document = reader.read(f.absolutePath)
    def ppt = new XMLSlideShow();

    def master = ppt.getSlideMasters().get(0)
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
    createToc(ppt, master, document)



    document.blocks.eachWithIndex { chapter, index ->
      if (chapter.title != null) {
        // toc
        createToc(ppt, master, document)
        // == section
        def sectionSlide = ppt.createSlide(master.getLayout("Section Header"))
        //sectionSlide.getPlaceholders().each{it.clearText()}
        def sectionHeader = sectionSlide.getPlaceholder(0)
        sectionHeader.setText(chapter.title)
      }

      createContents(chapter, ppt, master, f)
    }


    FileOutputStream out = new FileOutputStream(outputPath);
    ppt.write(out);
    ppt.close()
    out.close();
    sleep(100)
  }

  def beforeTitle = ""
  private void createContents(AbstractBlock chapter, XMLSlideShow ppt, master, File inputFile) {
    beforeTitle = chapter.title ?: beforeTitle
    def level = chapter.level
    chapter.blocks.eachWithIndex { i, idx ->
      if (i.level == level) {
        def slide = ppt.createSlide(master.getLayout("Title and Content"))
        slide.getPlaceholders().each { it.clearText() }
        def header = slide.getPlaceholder(0)
        header.setText(chapter.title ?: i.title ?: beforeTitle ?: "")
        // content or ===
        def content = slide.getPlaceholder(1)
        content.setText("")
        switch (i.class) {
          case ListImpl:
            appender(content, i.items, i.getContext().contains("olist"), 0)
            break
          case ListItemImpl:
            appenderItem(content, (ListItemImpl)i, i.getContext().contains("olist"), 0)
            break
          case BlockImpl:
            def b = i as BlockImpl
            switch (b.getBlockname()) {
              case "listing":
                def p = content.addNewTextParagraph()
                p.setBullet(false)
                def text = p.addNewTextRun()
                text.setText(i.content.toString())
                break
              case "paragraph":
                switch (b.content.toString()) {
                  case { it.startsWith("<span class=\"image\"><img ") }:
                    content.setText("")
                    def xml = Jsoup.parse(b.content.toString())
                    def img = xml.select("img")
                    def imgPath = img.attr("src")
                    byte[] pictureData = IOUtils.toByteArray(new FileInputStream(inputFile.parent + "/" + imgPath));
                    def pd = ppt.addPicture(pictureData, PictureData.PictureType.PNG);
                    def pic = slide.createPicture(pd);
                    break
                  case { it.startsWith("<a href=\"") }:
                    def xml = Jsoup.parse(b.content.toString())
                    def link = xml.select("a")
                    def linkPath = link.attr("href")
                    def text = link.text()
                    def p = content.addNewTextParagraph()
                    def t = p.addNewTextRun()
                    t.setText(text)
                    def hyperLink = t.createHyperlink()
                    hyperLink.linkToUrl(linkPath)
                    break
                  default:
                    content.appendText(b.content.toString(), true)
                }
                break
              case "open":
                println "plantuml ${b.getBlockname()} not suppert"
                content.setText("")
                break
              case "admonition":
                println "マーカー ${b.getBlockname()} not suppert"
                content.setText("")
                break
              default:
                println "other ${b.getBlockname()} not suppert"
                content.setText("")
                break
            }
            break
          default:
            content.appendText(i.content.toString(), true)
            break
        }
      } else {
        println i.class
        println i.content
        createContents(i, ppt, master, inputFile)
      }
    }
  }

  private void createToc(XMLSlideShow ppt, XSLFSlideMaster master, Document document) {
    def toc = ppt.createSlide(master.getLayout("Title and Content"))
    toc.getPlaceholders().each { it.clearText() }
    def tocHeader = toc.getPlaceholder(0)
    tocHeader.setText("Agenda")
    document.blocks.eachWithIndex { chapter, index ->
      def content = toc.getPlaceholder(1)
      content.appendText(chapter.title, true)
    }
  }


  void generate(String inputPath) {
    //new File("generated.pptx").deleteOnExit()
    generate(inputPath, "generated.pptx")
  }
}
