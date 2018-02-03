package org.gnote.gateway

import org.asciidoctor.Asciidoctor
import org.asciidoctor.ast.Document

/**
 * Created by kyon_mm on 2018/01/17.
 */
class AsciidocReader {

  Document read(String filePath) {
    def asciidoctor = Asciidoctor.Factory.create()
    return asciidoctor.loadFile(new File(filePath), [:])
  }
}
