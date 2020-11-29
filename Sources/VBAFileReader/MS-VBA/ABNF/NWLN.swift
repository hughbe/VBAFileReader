//
//  NWLN.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//

/// [MS-OVBA] 2.1.1.10 NWLN
/// Specifies a new line.
/// ABNF syntax:
/// NWLN = (CR LF) / (LF CR)
public struct NWLN: ABNFGrammar {
    init(stream: inout ABNFStream) throws {
        let string = try stream.readString(count: 2)
        guard string == "\r\n" || string == "\n\r" else {
            throw VBAFileError.corrupted
        }
    }
}
