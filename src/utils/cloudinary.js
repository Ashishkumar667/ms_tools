const cloudinary = require("cloudinary").v2;
const { Readable } = require("stream");
require("dotenv").config();

/**
 * Cloudinary utility functions for uploading videos/recordings
 */

// Configure Cloudinary from environment variables
cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key: process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET,
});

/**
 * Upload video buffer to Cloudinary
 * @param {Buffer} videoBuffer - Video file buffer
 * @param {object} options - Upload options
 * @param {string} options.publicId - Public ID for the video (optional)
 * @param {string} options.folder - Folder path in Cloudinary (optional)
 * @param {object} options.resourceType - Resource type (video, image, etc.) - defaults to 'video'
 * @returns {Promise<object>} Cloudinary upload result with secure_url
 */
async function uploadVideoToCloudinary(videoBuffer, options = {}) {
  const {
    publicId,
    folder = "teams-recordings",
    resourceType = "video",
  } = options;

  return new Promise((resolve, reject) => {
    // Convert buffer to stream
    const stream = cloudinary.uploader.upload_stream(
      {
        resource_type: resourceType,
        folder: folder,
        public_id: publicId,
        format: "mp4",
        // Video optimization settings
        eager: [
          { quality: "auto", format: "mp4" },
        ],
        eager_async: true,
      },
      (error, result) => {
        if (error) {
          reject(error);
        } else {
          resolve(result);
        }
      }
    );

    // Convert buffer to readable stream and pipe to Cloudinary
    const bufferStream = Readable.from(videoBuffer);
    bufferStream.pipe(stream);
  });
}

/**
 * Upload video from URL directly to Cloudinary (more efficient for large files)
 * Note: Cloudinary doesn't support custom headers for authenticated URLs in upload.
 * For authenticated URLs, we need to download first, then upload the buffer.
 * @param {string} videoUrl - URL of the video to upload
 * @param {object} options - Upload options
 * @returns {Promise<object>} Cloudinary upload result with secure_url
 */
async function uploadVideoFromUrlToCloudinary(videoUrl, options = {}) {
  const {
    publicId,
    folder = "teams-recordings",
    resourceType = "video",
    headers = {}, // Note: Cloudinary doesn't support custom headers for authenticated URLs
  } = options;

  // If we have auth headers, we can't use direct URL upload
  // This will throw an error that the caller can catch and fall back to buffer upload
  if (headers && Object.keys(headers).length > 0) {
    throw new Error("Authenticated URLs require buffer upload. Use uploadVideoToCloudinary instead.");
  }

  return cloudinary.uploader.upload(videoUrl, {
    resource_type: resourceType,
    folder: folder,
    public_id: publicId,
    format: "mp4",
    // Video optimization settings
    eager: [
      { quality: "auto", format: "mp4" },
    ],
    eager_async: true,
  });
}

/**
 * Delete video from Cloudinary
 * @param {string} publicId - Public ID of the video
 * @param {string} resourceType - Resource type (default: 'video')
 * @returns {Promise<object>} Deletion result
 */
async function deleteVideoFromCloudinary(publicId, resourceType = "video") {
  return cloudinary.uploader.destroy(publicId, {
    resource_type: resourceType,
  });
}

module.exports = {
  uploadVideoToCloudinary,
  uploadVideoFromUrlToCloudinary,
  deleteVideoFromCloudinary,
  cloudinary,
};

