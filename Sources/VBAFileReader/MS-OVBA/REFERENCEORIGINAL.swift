//
//  REFERENCEORIGINAL.swift
//
//
//  Created by Hugh Bellamy on 25/11/2020.
//

import DataStream

/// [MS-OVBA] 2.3.4.2.2.4 REFERENCEORIGINAL Record
/// Specifies the identifier of an Automation type library and a related REFERENCECONTROL (section 2.3.4.2.2.3).
/// The contained REFERENCECONTROL's twiddled type library was generated from the Automation type library identifier specified.
public struct REFERENCEORIGINAL {
    public let id: UInt16
    public let sizeOfLibidOriginal: UInt32
    public let libidOriginal: String
    public let referenceRecord: REFERENCECONTROL
    
    public init(dataStream: inout DataStream) throws {
        /// Id (2 bytes): An unsigned integer that specifies the identifier for this record. MUST be 0x0033.
        self.id = try dataStream.read(endianess: .littleEndian)
        guard self.id == 0x0033 else {
            throw VBAFileError.corrupted
        }
        
        /// SizeOfLibidOriginal (4 bytes): An unsigned integer that specifies the size in bytes of LibidOriginal.
        self.sizeOfLibidOriginal = try dataStream.read(endianess: .littleEndian)
        
        /// LibidOriginal (variable): An array of SizeOfLibidOriginal bytes that specifies the identifier of the Automation type library a
        /// REFERENCECONTROL (section 2.3.4.2.2.3) was generated from. MUST contain MBCS characters encoded using the code page
        /// specified in PROJECTCODEPAGE (section 2.3.4.2.1.4). MUST NOT contain null characters. MUST conform to the ABNF grammar in
        /// LibidReference (section 2.1.1.8).
        self.libidOriginal = try dataStream.readString(count: Int(self.sizeOfLibidOriginal), encoding: .ascii)!
        
        /// ReferenceRecord (variable): A ReferenceRecord of type REFERENCECONTROL (section 2.3.4.2.2.3) that specifies the type library
        /// generated from the Automation type library indicated by LibidOriginal.
        self.referenceRecord = try REFERENCECONTROL(dataStream: &dataStream)
    }
}
