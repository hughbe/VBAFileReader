//
//  WSP.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//

/// [MS-OVBA] 2.1.1.1 Common ABNF Rules
/// The following ABNF rules are used by section 2 and are included for reference. For more information, see [RFC4234] Appendix B.
/// ABNF Syntax:
/// CR = %x0D
/// DIGIT = %x30-39
/// DQUOTE = %x22
/// HEXDIG = DIGIT / "A" / "B" / "C" / "D" / "E" / "F"
/// HTAB = %x09
/// LF = %x0A
/// SP = %x20
/// VCHAR = %x21-7E
/// WSP = SP / HTAB
internal struct WSP: ABNFGrammar {
    init(stream: inout ABNFStream) throws {
        let char = try stream.readByte()
        guard char == 0x20 || char == 0x09 else {
            throw VBAFileError.corrupted
        }
    }
    
    public static func readMultiple(stream: inout ABNFStream) throws {
        while stream.position < stream.count {
            let nextByte = try stream.peekByte()
            guard nextByte == 0x20 || nextByte == 0x09 else {
                return
            }
            
            let _ = try stream.readByte()
        }
    }
}
