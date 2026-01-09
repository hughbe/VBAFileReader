//
//  PATH.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//

/// [MS-OVBA] 2.1.1.11 PATH
/// An array of characters that specifies a path to a file. MUST be less than 260 characters.
/// ABNF syntax:
/// PATH = DQUOTE *259QUOTEDCHAR DQUOTE
public struct PATH: ABNFGrammar {
    public let value: String

    init(stream: inout ABNFStream) throws {
        try stream.require(DQUOTE.self)
        self.value = try QUOTEDCHAR.read(stream: &stream)
        guard self.value.count < 260 else {
            throw VBAFileError.corrupted
        }
        try stream.require(DQUOTE.self)
    }
}
