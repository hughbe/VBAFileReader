//
//  MODULEOFFSET.swift
//
//
//  Created by Hugh Bellamy on 25/11/2020.
//

import DataStream

/// [MS-OVBA] 2.3.4.2.3.2.5 MODULEOFFSET Record
/// Specifies the location of the source code within the ModuleStream (section 2.3.4.3) that corresponds to the containing MODULE Record
/// (section 2.3.4.2.3.2).
public struct MODULEOFFSET {
    public let id: UInt16
    public let size: UInt32
    public let textOffset: UInt32
    
    public init(dataStream: inout DataStream) throws {
        /// Id (2 bytes): An unsigned integer that specifies the identifier for this record. MUST be 0x0031.
        self.id = try dataStream.read(endianess: .littleEndian)
        guard self.id == 0x0031 else {
            throw VBAFileError.corrupted
        }
        
        /// Size (4 bytes): An unsigned integer that specifies the size of TextOffset. MUST be 0x00000004.
        self.size = try dataStream.read(endianess: .littleEndian)
        guard self.size == 0x00000004 else {
            throw VBAFileError.corrupted
        }
        
        /// TextOffset (4 bytes): An unsigned integer that specifies the byte offset of the source code in the ModuleStream (section 2.3.4.3) named by
        /// MODULESTREAMNAME Record (section 2.3.4.2.3.2.3).
        self.textOffset = try dataStream.read(endianess: .littleEndian)
    }
}
