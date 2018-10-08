package com.vmware.jct.common.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import com.amazonaws.AmazonClientException;
import com.amazonaws.AmazonServiceException;
import com.amazonaws.auth.BasicAWSCredentials;
import com.amazonaws.auth.EnvironmentVariableCredentialsProvider;
import com.amazonaws.services.s3.AmazonS3;
import com.amazonaws.services.s3.AmazonS3Client;
import com.amazonaws.services.s3.model.CannedAccessControlList;
import com.amazonaws.services.s3.model.ObjectMetadata;
import com.amazonaws.services.s3.model.PutObjectRequest;

public class AwsTest {

	//http://jobcrafting.s3-website-us-west-2.amazonaws.com/images/profile1.jpg
	private static String bucketName = "jobcrafting";
	private static String keyName = "images/profile1.jpg";
	private static String uploadFileName = "/Users/rajan/Documents/tmp/img.JPG";
	private static final String AWS_ACCESS_KEY_ID = "CHANGE_ME";
	private static final String AWS_SECRET_ACCESS_KEY = "CHANGE_ME";
	
	
	public static void main(String[] args) throws IOException {
		AmazonS3 s3client = new AmazonS3Client(new BasicAWSCredentials(AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY));
		try {
			System.out.println("Uploading a new object to S3 from a file\n");
			File file = new File(uploadFileName);
			ObjectMetadata metadata = new ObjectMetadata();
			metadata.setContentType("image/jpeg");
			PutObjectRequest por = new PutObjectRequest(bucketName, keyName, new FileInputStream(file), metadata);
			por.setCannedAcl(CannedAccessControlList.PublicRead);
			s3client.putObject(por);
			//s3client.putObject(bucketName, keyName, new FileInputStream(file), metadata);

		} catch (AmazonServiceException ase) {
			System.out.println("Caught an AmazonServiceException, which " + "means your request made it "
					+ "to Amazon S3, but was rejected with an error response" + " for some reason.");
			System.out.println("Error Message:    " + ase.getMessage());
			System.out.println("HTTP Status Code: " + ase.getStatusCode());
			System.out.println("AWS Error Code:   " + ase.getErrorCode());
			System.out.println("Error Type:       " + ase.getErrorType());
			System.out.println("Request ID:       " + ase.getRequestId());
		} catch (AmazonClientException ace) {
			System.out.println("Caught an AmazonClientException, which " + "means the client encountered "
					+ "an internal error while trying to " + "communicate with S3, "
					+ "such as not being able to access the network.");
			System.out.println("Error Message: " + ace.getMessage());
		}
	}
}
