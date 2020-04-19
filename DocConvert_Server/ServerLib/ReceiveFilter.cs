﻿using SuperSocket.Common;
using SuperSocket.Facility.Protocol;
using SuperSocket.SocketBase.Protocol;
using System;

namespace DocConvert_Server
{
    public class EFBinaryRequestInfo : BinaryRequestInfo
    {
        public Int32 PacketID { get; private set; }
        public Int16 Value1 { get; private set; }
        public Int16 Value2 { get; private set; }

        public EFBinaryRequestInfo(int packetID, short value1, short value2, byte[] body)
            : base(null, body)
        {
            this.PacketID = packetID;
            this.Value1 = value1;
            this.Value2 = value2;
        }
    }

    public class ReceiveFilter : FixedHeaderReceiveFilter<EFBinaryRequestInfo>
    {
        public ReceiveFilter()
            : base(12)
        {

        }

        protected override int GetBodyLengthFromHeader(byte[] header, int offset, int length)
        {
            if (!BitConverter.IsLittleEndian)
                Array.Reverse(header, offset + 8, 4);

            var nBodySize = BitConverter.ToInt32(header, offset + 8);
            return nBodySize;
        }

        protected override EFBinaryRequestInfo ResolveRequestInfo(ArraySegment<byte> header, byte[] bodyBuffer, int offset, int length)
        {
            if (!BitConverter.IsLittleEndian)
                Array.Reverse(header.Array, 0, 12);

            return new EFBinaryRequestInfo(BitConverter.ToInt32(header.Array, 0),
                                           BitConverter.ToInt16(header.Array, 0 + 4),
                                           BitConverter.ToInt16(header.Array, 0 + 6),
                                           bodyBuffer.CloneRange(offset, length));
        }
    }
}
