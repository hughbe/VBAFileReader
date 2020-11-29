//
//  CompoundFileReader.swift
//  
//
//  Created by Hugh Bellamy on 25/11/2020.
//

import CompoundFileReader
import DataStream
import Foundation

/// [MS-OVBA] 2.2 File Structure
/// Specifies a VBA project and contained project items. All data is stored in a structured storage as specified in [MS-CFB]. The storages and streams
/// MUST be organized according to a hierarchy rooted at the Project Root Storage (section 2.2.1) as depicted in the following figure.
/// [MS-OVBA] 2.2.1 Project Root Storage
/// A single root storage. MUST contain VBA Storage (section 2.2.2) and PROJECT Stream (section 2.2.7). Optionally contains PROJECTwm Stream
/// (section 2.2.8), PROJECTlk Stream (section 2.2.9), and Designer Storages (section 2.2.10).
public struct VBAFile {
    public let vbaStorage: VBAStorage
    public let vbaProject: PROJECTStream
    public let projectwm: PROJECTwmStream?
    public let projectlk: PROJECTlkStream?
    public let designerModules: [String: DesignerModule]
    
    public init(data: Data) throws {
        let file = try CompoundFile(data: data)
        try self.init(storage: file.rootStorage)
    }
    
    public init(storage: CompoundFileStorage) throws {
        var storage = storage
        
        /// [MS-OVBA] 2.2.2 VBA Storage
        /// A storage that specifies VBA project and module information. MUST have the name "VBA" (caseinsensitive). MUST contain
        /// _VBA_PROJECT Stream (section 2.3.4.1) and dir Stream (section 2.3.4.2). MUST contain a Module Stream (section 2.2.5) for each module
        /// in the VBA project. Optionally contains SRP Streams (section 2.2.6).
        guard let vbaStorage = storage.children["VBA"] else {
            throw VBAFileError.corrupted
        }
        
        self.vbaStorage = try VBAStorage(storage: vbaStorage)
        
        /// [MS-OVBA] 2.2.7 PROJECT Stream
        /// A stream that specifies VBA project properties. MUST have the name "PROJECT" (case-insensitive). MUST contain data as specified
        /// by PROJECT Stream (section 2.3.1).
        guard let projectStorage = storage.children["PROJECT"] else {
            throw VBAFileError.corrupted
        }
        
        var projectDataStream = projectStorage.dataStream
        self.vbaProject = try PROJECTStream(dataStream: &projectDataStream, count: projectDataStream.count)
        
        /// [MS-OVBA] 2.2.8 PROJECTwm Stream
        /// A stream that specifies names of modules represented in both MBCS and UTF-16 encoding. MUST have the name
        /// "PROJECTwm" (case-insensitive). MUST contain data as specified by PROJECTwm Stream (section 2.3.3).
        if let projectwmStorage = storage.children["PROJECTwm"] {
            var projectwmDataStream = projectwmStorage.dataStream
            self.projectwm = try PROJECTwmStream(dataStream: &projectwmDataStream, count: projectwmDataStream.count)
        } else {
            self.projectwm = nil
        }
        
        /// [MS-OVBA] 2.2.9 PROJECTlk Stream
        /// A stream that specifies license information for ActiveX controls used in the VBA project. MUST have the name "PROJECTlk"
        /// (case-insensitive). MUST contain data as specified by PROJECTlk Stream (section 2.3.2).
        if let projectlkStorage = storage.children["PROJECTlk"] {
            var projectlkDataStream = projectlkStorage.dataStream
            self.projectlk = try PROJECTlkStream(dataStream: &projectlkDataStream)
        } else {
            self.projectlk = nil
        }
        
        /// [MS-OVBA] 2.2.10 Designer Storages
        /// A designer storage MUST be present for each designer module in the VBA project. The name is specified by MODULESTREAMNAME
        /// (section 2.3.4.2.3.2.3). MUST contain VBFrame Stream (section 2.3.5). If the designer is an Office Form ActiveX control, then this
        /// storage MUST contain storages and streams as specified by [MS-OFORMS] section 2.
        var designerModules: [String: DesignerModule] = [:]
        for module in self.vbaStorage.dir.modulesRecord.modules {
            guard let designerStorage = storage.children[module.nameRecord.moduleName] else {
                continue
            }
            
            designerModules[module.nameRecord.moduleName] = try DesignerModule(storage: designerStorage)
        }
        
        self.designerModules = designerModules
    }
}
