//
//  QUOTEDCHAR.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//

/// [MS-OVBA] 2.1.1.13 QUOTEDCHAR
/// Specifies a single character.
/// ABNF syntax:
/// QUOTEDCHAR = WSP / NQCHAR / ( DQUOTE DQUOTE )
/// NQCHAR = %x21 / %x23-FF
/// <DQUOTE DQUOTE>: Specifies a single double-quotation (") character.
public struct QUOTEDCHAR: ABNFGrammar {
    public let value: String
    
    init(stream: inout ABNFStream) throws {
        if stream.peek(string: "\"\"") {
            self.value = "\""
        } else {
            self.value = try stream.readString(count: 1)
        }
    }
    
    static func read(stream: inout ABNFStream) throws -> String {
        var value = ""
        while stream.position < stream.count {
            let startPosition = stream.position
            if try stream.readByte() == "\"".first!.asciiValue && stream.position < stream.count {
                if try stream.readByte() != "\"".first!.asciiValue {
                    stream.position = startPosition
                    break
                }
            }
            stream.position = startPosition
            
            value += try stream.read(QUOTEDCHAR.self).value
        }
        
        return value
    }
}
