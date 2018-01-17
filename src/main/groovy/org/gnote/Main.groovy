package org.gnote

import org.gnote.gateway.AsciidocReader
import org.gnote.gateway.PowerPointWriter

class Main {
    static void main(String[] args){
        println args
        def ppt = new PowerPointWriter(new AsciidocReader())
        ppt.generate(args[0])
    }
}
