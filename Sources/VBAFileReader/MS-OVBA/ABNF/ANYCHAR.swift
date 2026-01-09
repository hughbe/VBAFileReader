//
//  ANYCHAR.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//

/// [MS-OVBA] 2.1.1.2 ANYCHAR
/// Specifies any character value that is not a carriage-return, line-feed, or null.
/// ABNF syntax:
/// ANYCHAR = %x01-09 / %x0B / %x0C / %x0E-FF
public struct ANYCHAR: ABNFGrammar {
    public let value: String
    
    init (stream: inout ABNFStream) throws {
        let byte = try stream.readByte()
        guard (byte >= 0x01 && byte <= 0x09) || byte == 0x0B || byte == 0x0C || (byte >= 0x0E && byte <= 0xFF) else {
            throw VBAFileError.corrupted
        }
        
        self.value = try stream.readString(count: 1)
    }
    
    static func read(stream: inout ABNFStream) throws -> String {
        var s = ""
        while stream.position < stream.count {
            let byte = try stream.peekByte()
            guard (byte >= 0x01 && byte <= 0x09) || byte == 0x0B || byte == 0x0C || (byte >= 0x0E && byte <= 0xFF) else {
                break
            }
            
            s += try stream.read(ANYCHAR.self).value
        }
        
        return s
    }
}
