using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.Services
{
    public class StreamProxy : Stream
    {
		readonly Stream _original;
		public StreamProxy(Stream original)
		{
			_original = original;
		}

		public override bool CanRead => _original.CanRead;

		public override bool CanSeek => _original.CanSeek;

		public override bool CanWrite => _original.CanWrite;

		public override long Length => _original.Length;

		public override long Position { get => _original.Position; set => _original.Position = value; }

		public override void Flush()
		{
			_original.Flush();
		}

		public override int Read(byte[] buffer, int offset, int count)
		{
			return _original.Read(buffer, offset, count);
		}

		public override long Seek(long offset, SeekOrigin origin)
		{
			return _original.Seek(offset, origin);
		}

		public override void SetLength(long value)
		{
			_original.SetLength(value);
		}

        public override int Read(Span<byte> buffer)
        {
            return _original.Read(buffer);
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            return _original.ReadAsync(buffer, offset, count, cancellationToken);
        }

        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            return _original.ReadAsync(buffer, cancellationToken);
        }

        public override IAsyncResult BeginRead(byte[] buffer, int offset, int count, AsyncCallback callback, object state)
        {
            return _original.BeginRead(buffer, offset, count, callback, state);
        }

        public override Task CopyToAsync(Stream destination, int bufferSize, CancellationToken cancellationToken)
        { 
            return _original.CopyToAsync(destination, bufferSize, cancellationToken);
        }

        public override void CopyTo(Stream destination, int bufferSize)
        {
			_original.CopyTo(destination, bufferSize);
        }

        public override int EndRead(IAsyncResult asyncResult)
        {
            return _original.EndRead(asyncResult);
        }

        public override int ReadByte()
        {
            return _original.ReadByte();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            _original.Write(buffer, offset, count);
        }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            _original.Write(buffer);
        }

        public override IAsyncResult BeginWrite(byte[] buffer, int offset, int count, AsyncCallback callback, object state)
        {
            return _original.BeginWrite(buffer, offset, count, callback, state);
        }

        public override void EndWrite(IAsyncResult asyncResult)
        {
            _original.EndWrite(asyncResult);
        }

        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            return _original.WriteAsync(buffer, offset, count, cancellationToken);
        }

        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
        {
            return _original.WriteAsync(buffer, cancellationToken);
        }

        public override Task FlushAsync(CancellationToken cancellationToken)
        {
            return _original.FlushAsync(cancellationToken);
        }

        public override void WriteByte(byte value)
        {
            _original.WriteByte(value);
        }
    }
}
