package com.inventory.manager.utils;

import android.content.Context;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.net.Uri;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

public class ImageUtils {

    private static final int MAX_IMAGE_SIZE = 800;

    public static File createImageFile(Context context) throws IOException {
        String timeStamp = new SimpleDateFormat("yyyyMMdd_HHmmss", Locale.getDefault()).format(new Date());
        String imageFileName = "ITEM_" + timeStamp;
        File storageDir = new File(context.getFilesDir(), "images");
        if (!storageDir.exists()) {
            storageDir.mkdirs();
        }
        return File.createTempFile(imageFileName, ".jpg", storageDir);
    }

    public static String saveImageFromUri(Context context, Uri uri) {
        try {
            InputStream inputStream = context.getContentResolver().openInputStream(uri);
            if (inputStream == null) return null;

            Bitmap bitmap = BitmapFactory.decodeStream(inputStream);
            inputStream.close();

            if (bitmap == null) return null;

            Bitmap resized = resizeBitmap(bitmap, MAX_IMAGE_SIZE);
            if (resized != bitmap) {
                bitmap.recycle();
            }
            bitmap = resized;

            File imageFile = createImageFile(context);
            FileOutputStream fos = new FileOutputStream(imageFile);
            bitmap.compress(Bitmap.CompressFormat.JPEG, 85, fos);
            fos.flush();
            fos.close();
            bitmap.recycle();

            return imageFile.getAbsolutePath();
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }

    public static String saveImageFromFile(Context context, String filePath) {
        try {
            if (filePath == null) return null;

            Bitmap bitmap = BitmapFactory.decodeFile(filePath);
            if (bitmap == null) return null;

            Bitmap resized = resizeBitmap(bitmap, MAX_IMAGE_SIZE);
            if (resized != bitmap) {
                bitmap.recycle();
            }

            File imageFile = createImageFile(context);
            FileOutputStream fos = new FileOutputStream(imageFile);
            resized.compress(Bitmap.CompressFormat.JPEG, 85, fos);
            fos.flush();
            fos.close();
            resized.recycle();

            // Delete original full-resolution file if a new file was created
            String newPath = imageFile.getAbsolutePath();
            if (!newPath.equals(filePath)) {
                new File(filePath).delete();
            }

            return newPath;
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }

    private static Bitmap resizeBitmap(Bitmap bitmap, int maxSize) {
        int width = bitmap.getWidth();
        int height = bitmap.getHeight();

        if (width <= maxSize && height <= maxSize) {
            return bitmap;
        }

        float ratio = Math.min((float) maxSize / width, (float) maxSize / height);
        int newWidth = Math.round(width * ratio);
        int newHeight = Math.round(height * ratio);

        return Bitmap.createScaledBitmap(bitmap, newWidth, newHeight, true);
    }

    public static void deleteImage(String imagePath) {
        if (imagePath != null) {
            File file = new File(imagePath);
            if (file.exists()) {
                file.delete();
            }
        }
    }
}
