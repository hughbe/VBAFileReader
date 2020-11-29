//
//  ModuleIdentifier.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//


/// [MS-OVBA] 2.1.1.9 ModuleIdentifier
/// Specifies the name of a module. SHOULD be an identifier as specified by [MS-VBAL] section 3.3.5. MAY<2> be any string of
/// characters. MUST be less than or equal to 31 characters long.
public struct ModuleIdentifier: ABNFGrammar {
    public let value: String
    
    init(stream: inout ABNFStream) throws {
        let position = stream.position
        var count = 0
        while stream.position < stream.count {
            let byte = try stream.readByte()
            guard byte != "\n".first!.asciiValue &&
                    byte != "\r".first!.asciiValue &&
                    byte != "/".first!.asciiValue &&
                    byte != "=".first!.asciiValue &&
                    byte != 0x09 /* HTAB */ &&
                    byte != 0x0A /* LF */ &&
                    byte != 0x20 /* SP */ else {
                break
            }

            count += 1
        }
        
        stream.position = position
        self.value = try stream.readString(count: count)
    }
}
