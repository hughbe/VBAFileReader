//
//  DesignerModule.swift
//  
//
//  Created by Hugh Bellamy on 28/11/2020.
//

import CompoundFileReader

/// [MS-OVBA] 2.2.10 Designer Storages
/// A designer storage MUST be present for each designer module in the VBA project. The name is specified by MODULESTREAMNAME
/// (section 2.3.4.2.3.2.3). MUST contain VBFrame Stream (section 2.3.5). If the designer is an Office Form ActiveX control, then this storage
/// MUST contain storages and streams as specified by [MS-OFORMS] section 2.
public struct DesignerModule {
    public let frame: VBFrame
    public let storage: CompoundFileStorage
    
    public init(storage: CompoundFileStorage) throws {
        var storage = storage

        /// [MS-OVBA] 2.2.11 VBFrame Stream
        /// A stream that specifies designer module properties. MUST contain data as specified by VBFrame Stream (section 2.3.5). Name of
        /// this stream MUST start with the UTF-16 character 0x0003 followed by the UTF-16 string "VBFrame" (case-insensitive).
        guard let frameStorage = storage.children["\u{0003}VBFrame"] else {
            throw VBAFileError.corrupted
        }
        
        var frameDataStream = frameStorage.dataStream
        self.frame = try VBFrame(dataStream: &frameDataStream, count: frameDataStream.count)
        
        self.storage = storage
    }
}
