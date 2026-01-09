//
//  VBABOOL.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//

/// [MS-OVBA] 2.1.1.14 VBABOOL
/// Specifies a Boolean value.
/// Value Meaning
/// "0" FALSE
/// "-1" TRUE
/// ABNF syntax:
/// VBABOOL = "0" / "-1"
public struct VBABOOL: ABNFGrammar {
    public let value: Bool
    
    init(stream: inout ABNFStream) throws {
        if stream.peek(string: "0") {
            try stream.require("0")
            self.value = false
        } else if stream.peek(string: "-1") {
            try stream.require("-1")
            self.value = true
        } else {
            throw VBAFileError.corrupted
        }
    }
}

