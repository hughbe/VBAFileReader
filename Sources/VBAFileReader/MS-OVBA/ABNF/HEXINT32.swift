//
//  HEXINT32.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//

/// [MS-OVBA] 2.1.1.6 HEXINT32
/// Specifies a hexadecimal-encoded signed integer. MUST be between −2147483648 and 2147483647.
/// ABNF syntax:
/// HEXINT32 = "&H" 8HEXDIG
public struct HEXINT32: ABNFGrammar {
    public let value: Int32
    
    init(stream: inout ABNFStream) throws {
        try stream.require("&H")
        let stringValue = try stream.readString(count: 8)
        guard let value = Int32(stringValue, radix: 16) else {
            throw VBAFileError.corrupted
        }
        
        self.value = value
    }
}
