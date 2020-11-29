//
//  VBAStorage.swift
//  
//
//  Created by Hugh Bellamy on 25/11/2020.
//

import CompoundFileReader
import DataStream

/// [MS-OVBA] 2.2.2 VBA Storage
/// A storage that specifies VBA project and module information. MUST have the name "VBA" (caseinsensitive). MUST contain _VBA_PROJECT
/// Stream (section 2.3.4.1) and dir Stream (section 2.3.4.2). MUST contain a Module Stream (section 2.2.5) for each module in the VBA project.
/// Optionally contains SRP Streams (section 2.2.6).
public struct VBAStorage {
    private let storage: CompoundFileStorage
    public let project: _VBA_PROJECT
    public let dir: DirStream
    public let modules: [String: ModuleStream]
    public let srpStreams: [CompoundFileStorage]
    
    public init(storage: CompoundFileStorage) throws {
        var storage = storage
        
        /// [MS-OVBA] 2.2.3 _VBA_PROJECT Stream
        /// A stream that specifies the version-dependent project information. MUST have the name "_VBA_PROJECT" (case-insensitive). MUST
        /// contain data as specified by _VBA_PROJECT Stream (section 2.3.4.1).
        guard let projectStorage = storage.children["_VBA_PROJECT"] else {
            throw VBAFileError.corrupted
        }
        
        var projectDataStream = projectStorage.dataStream
        self.project = try _VBA_PROJECT(dataStream: &projectDataStream, count: projectDataStream.count)
        
        /// [MS-OVBA] 2.2.4 dir Stream
        /// A stream that specifies VBA project properties, project references, and module properties. MUST have the name "dir" (case-insensitive).
        /// MUST contain data as specified by dir Stream (section 2.3.4.2).
        guard let dirStorage = storage.children["dir"] else {
            throw VBAFileError.corrupted
        }
        
        var dirDataStreamCompressed = dirStorage.dataStream
        var dirDataStream = DataStream(try VBACompression.decompress(dataStream: &dirDataStreamCompressed))
        self.dir = try DirStream(dataStream: &dirDataStream)
        
        /// [MS-OVBA] 2.2.5 Module Stream
        /// A stream that specifies the source code of modules in the VBA project. The name of this stream is specified by MODULESTREAMNAME
        /// (section 2.3.4.2.3.2.3). MUST contain data as specified by Module Stream (section 2.3.4.3).
        var modules: [String: ModuleStream] = [:]
        modules.reserveCapacity(Int(self.dir.modulesRecord.count))
        for module in self.dir.modulesRecord.modules {
            guard let storage = storage.children[module.nameRecord.moduleName] else {
                continue
            }
            
            var dataStream = storage.dataStream
            let moduleStream = try ModuleStream(dataStream: &dataStream,
                                                offset: Int(module.offsetRecord.textOffset),
                                                count: dataStream.count)
            modules[module.nameRecord.moduleName] = moduleStream
        }
        
        self.modules = modules
        
        /// [MS-OVBA] 2.2.6 SRP Streams
        /// Streams that specify an implementation-specific and version-dependent performance cache. MUST be ignored on read.
        /// MUST NOT be present on write.
        /// The name of each of these streams is specified by the following ABNF grammar:
        /// SRPStreamName = "__SRP_" 1*25DIGIT
        self.srpStreams = storage.children.values.filter { $0.name.hasPrefix("__SRP_") }.sorted { $0.name < $1.name }
        
        self.storage = storage
    }
}
