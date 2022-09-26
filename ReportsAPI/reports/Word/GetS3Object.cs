using System;
using System.IO;
using Amazon.S3;
using Amazon.S3.Model;
using Amazon;
using reports.Services.ConfigurationProviders;
using reports.Extensions;

namespace reports.Word
{
    
    public class GetS3Object
    {

        private const string bucketName = "buildee-test";
        private static readonly RegionEndpoint bucketRegion = RegionEndpoint.USWest2;
        private static IAmazonS3 client;

        private readonly IConfigProvider configProvider;

        public GetS3Object()
        {
            configProvider = new ConfigProvider();
        }

        public string ReadObjectData(string fileName)
        {

            // get the document from the S3 bucket
            client = new AmazonS3Client(configProvider.AWSAccessKey, configProvider.AWSAccessSecretKey, bucketRegion);
            string keyName = fileName.Decode();
            try
            {
                GetObjectRequest request = new GetObjectRequest
                {
                    BucketName = bucketName,
                    Key = keyName
                };
                GetObjectResponse response = client.GetObject(request);

                // store file on local machine
                string tmpFileName = $"{Path.GetTempFileName()}{Guid.NewGuid().ToString()}.docx";
                response.WriteResponseStreamToFile(tmpFileName);
                
                // return file path to downloaded s3 object
                return tmpFileName;
            }
            catch (AmazonS3Exception e)
            {
                return "AmazonS3 error" + e;
            }
            catch (Exception e)
            {
                return "Unknown error" + e;
            }
        }
    }
}