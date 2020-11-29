//
//  EQ.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//

/// [MS-OVBA] 2.1.1.3 EQ
/// Defines syntax for separating a property name from a value.
/// ABNF syntax:
/// EQ = *WSP "=" *WSP
public struct EQ: ABNFGrammar {
    init(stream: inout ABNFStream) throws {
        try WSP.readMultiple(stream: &stream)
        try stream.require("=")
        try WSP.readMultiple(stream: &stream)
    }
}
