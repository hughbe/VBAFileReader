//
//  DecompressedBuffer.swift
//  
//
//  Created by Hugh Bellamy on 25/11/2020.
//

import DataStream
import Foundation

/// [MS-OVBA] 2.4.1.1.2 DecompressedBuffer
/// The DecompressedBuffer is a resizable array of bytes that contains the same data as the CompressedContainer (section 2.4.1.1.1), but the data is in
/// an uncompressed format.
public struct DecompressedBuffer {
    public let decompressedChunks: [DecompressedChunk]
    
    public init(container: CompressedContainer) throws {
        var decompressedChunks: [DecompressedChunk] = []
        for compressedChunk in container.chunks {
            decompressedChunks.append(try DecompressedChunk(compressedChunk: compressedChunk))
        }
        
        self.decompressedChunks = decompressedChunks
    }
    
    public var data: Data {
        var result = Data()
        for chunk in decompressedChunks {
            var dataStream = chunk.dataStream
            result += try! dataStream.readBytes(count: chunk.dataStream.count)
        }
        
        return result
    }
}
