using Amazon;
using Amazon.S3;
using Amazon.S3.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using reports.Extensions;

namespace reports.Services.DocumentStorage
{
    public class S3DocumentStorage : IDocumentStorage
    {

        private readonly string _bucketName;
        private readonly AmazonS3Client _s3;

        public S3DocumentStorage(string accessKeyId, string secretAccessKey, string bucketName)
        {
            _s3 = new AmazonS3Client(accessKeyId, secretAccessKey, RegionEndpoint.USWest2);
            _bucketName = bucketName;
        }

        public byte[] Get(string id)
        {
            try
            {
                var document = _s3.GetObject(_bucketName, id);
                var bytes = document.ResponseStream.ReadToEnd();
                return bytes;
            }
            // TODO: add more error handling
            catch (AmazonS3Exception exception) when (exception.ErrorCode.Equals("NoSuchKey"))
            {
                var message = $"We're sorry but the document does not exist.";                
                throw exception;
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }

        public void Save(string id, byte[] bytes)
        {
            var stream = new MemoryStream(bytes);
            var request = new PutObjectRequest { BucketName = _bucketName, Key = id, InputStream = stream };

            _s3.PutObject(request);
        }

        public void Delete(string id)
        {
            _s3.DeleteObject(_bucketName, id);
        }
    }
}