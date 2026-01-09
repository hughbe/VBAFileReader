//
//  PROJECTCOMPATVERSION.swift
//  VBAFileReader
//
//  Created by Hugh Bellamy on 09/01/2026.
//

import DataStream

/// [MS-OVBA] 2.3.4.2.1.2 PROJECTCOMPATVERSION Record
/// Specifies the VBA project’s compat version.
public struct PROJECTCOMPATVERSION {
    public let id: UInt16
    public let size: UInt32
    public let compatVersion: UInt32
    
    
    public init(dataStream: inout DataStream) throws {
        /// Id (2 bytes): An unsigned integer that specifies the identifier for this record. MUST be 0x004A.
        self.id = try dataStream.read(endianess: .littleEndian)
        guard self.id == 0x004A else {
            throw VBAFileError.corrupted
        }
        
        /// Size (4 bytes): An unsigned integer that specifies the size of compat version. MUST be 0x00000004.
        self.size = try dataStream.read(endianess: .littleEndian)
        guard self.size == 0x00000004 else {
            throw VBAFileError.corrupted
        }
        
        /// CompatVersion (4 bytes):  An unsigned integer that specifies the compat version value for the VBA project.
        self.compatVersion = try dataStream.read(endianess: .littleEndian)
    }
}
