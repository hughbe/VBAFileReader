//
//  FLOAT.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//

/// [MS-OVBA] 2.1.1.4 FLOAT
/// Specifies a floating-point number.
/// ABNF syntax:
/// FLOAT = [SIGN] ( ( 1*DIGIT "." 1*DIGIT [EXP] ) /
///  ( "." 1*DIGIT [EXP] ) /
///  ( 1*DIGIT ["."] [EXP] ) )
/// EXP = "e" [SIGN] 1*DIGIT
/// SIGN = "+" / "-"
public struct FLOAT: ABNFGrammar {
    public let value: Float
    
    init(stream: inout ABNFStream) throws {
        var string = ""
        while stream.position < stream.count {
            if stream.peek(string: "+") || stream.peek(string: "-") {
                guard string.count == 0 || string.last == "e" else {
                    break
                }
            } else {
                let char = try stream.peekByte()
                guard char == ".".first!.asciiValue || char == "e".first!.asciiValue || char >= 0x30 && char <= 0x39 else {
                    break
                }
            }
            
            string.append(try stream.readString(count: 1))
        }
        
        guard let value = Float(string) else {
            throw VBAFileError.corrupted
        }
        
        self.value = value
    }
}

