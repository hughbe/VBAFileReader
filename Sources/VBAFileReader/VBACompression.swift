//
//  VBACompression.swift
//  
//
//  Created by Hugh Bellamy on 26/11/2020.
//

import DataStream
import Foundation

public struct VBACompression {
    public static func decompress(dataStream: inout DataStream) throws -> Data {
        /// [MS-OVBA] 2.4.1.3.1 Decompression Algorithm
        /// The Decompression algorithm pseudocode decodes the data in a CompressedContainer (section 2.4.1.1.1) and writes the
        /// uncompressed bytes to a DecompressedBuffer (section 2.4.1.1.2). The pseudocode first validates CompressedContainer
        /// SignatureByte (section 2.4.1.1.1). If validation fails, then the CompressedContainer (section 2.4.1.1.1) is corrupt and cannot be
        /// decoded. The pseudocode then iterates over the CompressedChunks (section 2.4.1.1.4). On each iteration, the current
        /// CompressedChunk is decoded.
        /// The pseudocode to decompress the CompressedContainer (section 2.4.1.1.1) into the DecompressedBuffer (section 2.4.1.1.2)
        /// uses the state variables described in State Variables (section 2.4.1.2): CompressedCurrent, CompressedRecordEnd, and
        /// DecompressedCurrent.
        /// These state variables MUST be initialized by the caller. CompressedChunkStart is also used.
        /// IF the byte located at CompressedCurrent EQUALS 0x01 THEN
        ///  INCREMENT CompressedCurrent
        ///  WHILE CompressedCurrent is LESS THAN CompressedRecordEnd
        ///  SET CompressedChunkStart TO CompressedCurrent
        ///  CALL Decompressing a CompressedChunk
        ///  END WHILE
        /// ELSE
        ///  RAISE ERROR
        /// ENDIF
        let container = try CompressedContainer(dataStream: &dataStream, count: dataStream.count)
        let decompressedBuffer = try DecompressedBuffer(container: container)
        return decompressedBuffer.data
    }
}
