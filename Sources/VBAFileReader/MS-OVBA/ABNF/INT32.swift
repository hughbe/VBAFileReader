//
//  INT32.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//

/// [MS-OVBA] 2.1.1.7 INT32
/// Specifies a signed integer. MUST be between −2147483648 and 2147483647. ABNF syntax:
/// INT32 = ["-"] 1*DIGIT
public struct INT32: ABNFGrammar {
    public let value: Int32
    
    init(stream: inout ABNFStream) throws {
        var string = ""
        while stream.position < stream.count {
            if stream.peek(string: "-") {
                guard string.count == 0 else {
                    break
                }
            } else {
                let char = try stream.peekByte()
                guard char >= 0x30 && char <= 0x39 else {
                    break
                }
            }
            
            string.append(try stream.readString(count: 1))
        }
        
        guard let value = Int32(string) else {
            throw VBAFileError.corrupted
        }
        
        self.value = value
    }
}
