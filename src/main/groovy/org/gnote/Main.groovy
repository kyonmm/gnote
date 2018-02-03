package org.gnote

import org.asciidoctor.Asciidoctor
import org.gnote.gateway.AsciidocReader
import org.gnote.gateway.PowerPointWriter

import static org.asciidoctor.OptionsBuilder.options

class Main {
  static void main(String[] args){
    println args
//        def asciidoctor = Asciidoctor.Factory.create();
//        asciidoctor.requireLibrary("asciidoctor-diagram");
//        def f = new File(args[0].replace(/~/, System.getProperty("user.home")))
//        asciidoctor.convertFile(f, options().backend("html5").get());
    def ppt = new PowerPointWriter(new AsciidocReader())
    ppt.generate(args[0])
  }
}
