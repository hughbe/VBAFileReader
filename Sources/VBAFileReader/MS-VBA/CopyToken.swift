//
//  CopyToken.swift
//  
//
//  Created by Hugh Bellamy on 25/11/2020.
//

import DataStream
import BitField

/// [MS-OVBA] 2.4.1.1.8 CopyToken
/// CopyToken is a two-byte record interpreted as an unsigned 16-bit integer in little-endian order. A CopyToken is a compressed encoding of an array of
/// bytes from a DecompressedChunk (section 2.4.1.1.3). The byte array encoded by a CopyToken is a byte-for-byte copy of a byte array elsewhere
/// in the same DecompressedChunk, called a CopySequence (section 2.4.1.3.19).
/// The starting location, in a DecompressedChunk, is determined by the Compressing a Token (section 2.4.1.3.9) and the Decompressing a Token
/// (section 2.4.1.3.5) algorithms. Packed into the CopyToken is the Offset, the distance, in byte count, to the beginning of the CopySequence. Also
/// packed into the CopyToken is the Length, the number of bytes encoded in the CopyToken. Length also specifies the count of bytes in the
/// CopySequence. The values encoded in Offset and Length are computed by the Matching (section 2.4.1.3.19.4) algorithm.
public struct CopyToken: Token {
    public let length: UInt16
    public let offset: UInt16
    
    public var count: Int { Int(length) }
    
    /// [MS-OVBA] 2.4.1.3.19.1 CopyToken Help
    /// CopyToken Help derived bit masks are used by the Unpack CopyToken (section 2.4.1.3.19.2) and the Pack CopyToken (section
    /// 2.4.1.3.19.3) algorithms. CopyToken Help also derives the maximum length for a CopySequence (section 2.4.1.3.19) which is used by
    /// the Matching algorithm (section 2.4.1.3.19.4).
    /// The pseudocode uses the state variables described in State Variables (section 2.4.1.2): DecompressedCurrent and
    /// DecompressedChunkStart.
    /// The pseudocode for CopyToken Help returns the following output parameters:
    /// LengthMask (2 bytes): An unsigned 16-bit integer. A bitmask used to access CopyToken.Length.
    /// OffsetMask (2 bytes): An unsigned 16-bit integer. A bitmask used to access CopyToken.Offset.
    /// BitCount (2 bytes): An unsigned 16-bit integer. The number of bits set to 0b1 in OffsetMask.
    /// MaximumLength (2 bytes): An unsigned 16-bit integer. The largest possible integral value that can fit into CopyToken.Length.
    private static func copyTokenHelp(difference: Int) -> (lengthMask: UInt16, offsetMask: UInt16, bitCount: UInt16, maximumLength: UInt16) {
        /// SET difference TO DecompressedCurrent MINUS DecompressedChunkStart
        /// SET BitCount TO the smallest integer that is GREATER THAN OR EQUAL TO LOGARITHM base 2 of difference
        var bitCount: UInt16 = 0
        while (1 << bitCount) < difference {
            bitCount += 1
        }
        
        /// SET BitCount TO the maximum of BitCount and 4
        bitCount = max(bitCount, 4)
        
        /// SET LengthMask TO 0xFFFF RIGHT SHIFT BY BitCount
        let lengthMask = UInt16(0xFFFF) >> bitCount
        
        /// SET OffsetMask TO BITWISE NOT LengthMask
        let offsetMask = ~lengthMask
        
        /// SET MaximumLength TO (0xFFFF RIGHT SHIFT BY BitCount) PLUS 3
        let maximumLength = (UInt16(0xFFFF) >> bitCount) + 3
        
        return (lengthMask: lengthMask,
                offsetMask: offsetMask,
                bitCount: bitCount,
                maximumLength: maximumLength)
    }
    
    /// [MS-OVBA] 2.4.1.3.19.2 Unpack CopyToken
    /// The Unpack CopyToken pseudocode will compute the specifications of a CopySequence (section 2.4.1.3.19) that are encoded in a
    /// CopyToken.
    /// The pseudocode for Unpack CopyToken takes the following input parameters:
    /// Token (2 bytes): A CopyToken (section 2.4.1.1.8).
    /// The pseudocode takes the following output parameters:
    /// Offset (2 bytes): An unsigned 16-bit integer that specifies the beginning of a CopySequence (section 2.4.1.3.19).
    /// Length (2 bytes): An unsigned 16-bit integer that specifies the length of a CopySequence (section 2.4.1.3.19) as follows:
    private static func unpack(token: UInt16, difference: Int) throws -> (offset: UInt16, length: UInt16) {
        /// 1. CALL CopyToken Help (section 2.4.1.3.19.1) returning LengthMask, OffsetMask, and BitCount.
        let (lengthMask, offsetMask, bitCount, maximumLength) = copyTokenHelp(difference: difference)
        
        /// 2. SET Length TO (Token BITWISE AND LengthMask) PLUS 3.
        let length = (token & lengthMask) + 3
        guard length <= maximumLength else {
            throw VBAFileError.corrupted
        }
        
        /// 3. SET temp1 TO Token BITWISE AND OffsetMask.
        let temp1 = token & offsetMask
        
        /// 4. SET temp2 TO 16 MINUS BitCount.
        let temp2 = 16 - bitCount
        
        /// 5. SET Offset TO (temp1 RIGHT SHIFT BY temp2) PLUS 1.
        let offset = (temp1 >> temp2) + 1
        
        return (offset: offset, length: length)
    }
    
    /// [MS-OVBA] 2.4.1.3.19.3 Pack CopyToken
    /// The Pack CopyToken pseudocode will take the Offset and Length values that specify a CopySequence (section 2.4.1.3.19) and pack
    /// them into a CopyToken (section 2.4.1.1.8).
    /// The Pack CopyToken pseudocode takes the following input parameters:
    /// Offset (2 bytes): An unsigned 16-bit integer that specifies the beginning of a CopySequence (section 2.4.1.3.19).
    /// Length (2 bytes): An unsigned 16-bit integer that specifies the length of a CopySequence (section 2.4.1.3.19).
    /// The Pack CopyToken pseudocode takes the following output parameters:
    /// Token (2 bytes): A CopyToken (section 2.4.1.1.8).
    private static func pack(offset: UInt16, length: UInt16, difference: Int) throws -> UInt16 {
        /// CALL CopyToken Help (section 2.4.1.3.19.1) returning LengthMask, OffsetMask, and BitCount
        let (_, _, bitCount, maximumLength) = copyTokenHelp(difference: difference)
        guard length <= maximumLength else {
            throw VBAFileError.corrupted
        }
        
        /// SET temp1 TO Offset MINUS 1
        let temp1 = offset - 1
        
        /// SET temp2 TO 16 MINUS BitCount
        let temp2 = 16 - bitCount
        
        /// SET temp3 TO Length MINUS 3
        let temp3 = length - 3
        
        /// SET Token TO (temp1 LEFT SHIFT BY temp2) BITWISE OR temp3
        return (temp1 << temp2) | temp3
    }
    
    public init(dataStream: inout DataStream, differenceFromChunkStart: Int) throws {
        let rawValue: UInt16 = try dataStream.read(endianess: .littleEndian)
        
        let (offset, length) = try CopyToken.unpack(token: rawValue, difference: differenceFromChunkStart)
        
        /// Length (variable): A variable bit unsigned integer that specifies the number of bytes contained in a CopySequence minus three. MUST be
        /// greater than or equal to zero. MUST be less than 4093. The number of bits used to encode Length MUST be greater than or equal to four.
        /// The number of bits used to encode Length MUST be less than or equal to 12. The number of bits used to encode Length is computed and
        /// used in the Unpack CopyToken (section 2.4.1.3.19.2) and the Pack CopyToken (section 2.4.1.3.19.3) algorithms.
        self.length = length
        
        /// Offset (variable): A variable bit unsigned integer that specifies the distance, in byte count, from the beginning of a duplicate set of bytes in
        /// the DecompressedBuffer to the beginning of a CopySequence. The value stored in Offset is the distance minus three. MUST be greater than
        /// zero. MUST be less than 4096. The number of bits used to encode Offset MUST be greater than or equal to four. The number of bits used
        /// to encode Offset MUST be less than or equal to 12. The number of bits used to encode Offset is computed and used in the Unpack
        /// CopyToken and the Pack CopyToken algorithms.
        self.offset = offset
    }
    
    public func decompress(to: inout [UInt8]) {
        let position = to.count - Int(offset)
        
        var copySequence: [UInt8] = [UInt8](repeating: 0, count: Int(length))
        for i in 0..<min(offset, length) {
            copySequence[Int(i)] = to[position + Int(i)]
        }
        
        if offset < length {
            for i in offset..<length {
                copySequence[Int(i)] = copySequence[Int(i) % Int(offset)]
            }
        }
        
        to.append(contentsOf: copySequence)
    }
}
