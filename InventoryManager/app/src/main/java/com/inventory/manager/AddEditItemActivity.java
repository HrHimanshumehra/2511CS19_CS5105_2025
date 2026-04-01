package com.inventory.manager;

import android.Manifest;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.graphics.Bitmap;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.provider.MediaStore;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.EditText;
import android.widget.ImageView;
import android.widget.Spinner;
import android.widget.Toast;

import androidx.annotation.NonNull;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;
import androidx.core.content.FileProvider;

import com.bumptech.glide.Glide;
import com.inventory.manager.database.AppDatabase;
import com.inventory.manager.database.ItemDao;
import com.inventory.manager.model.Item;
import com.inventory.manager.utils.Constants;
import com.inventory.manager.utils.ImageUtils;

import java.io.File;
import java.io.IOException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class AddEditItemActivity extends AppCompatActivity {

    private static final int REQUEST_CAMERA = 100;
    private static final int REQUEST_GALLERY = 101;
    private static final int PERMISSION_CAMERA = 200;
    private static final int PERMISSION_STORAGE = 201;

    private EditText etName, etQuantity, etPrice, etLowStockThreshold;
    private Spinner spinnerCategory;
    private ImageView ivItemImage;
    private Button btnCamera, btnGallery, btnSave;

    private ItemDao itemDao;
    private ExecutorService executor;
    private Item currentItem;
    private boolean isEditMode = false;
    private String currentImagePath;
    private String cameraPhotoPath;


    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_add_edit_item);

        setSupportActionBar(findViewById(R.id.toolbarAddEdit));
        if (getSupportActionBar() != null) {
            getSupportActionBar().setDisplayHomeAsUpEnabled(true);
        }

        executor = Executors.newSingleThreadExecutor();
        itemDao = AppDatabase.getInstance(this).itemDao();

        initViews();
        setupCategorySpinner();

        int itemId = getIntent().getIntExtra("item_id", -1);
        if (itemId != -1) {
            isEditMode = true;
            setTitle("Edit Item");
            loadItem(itemId);
        } else {
            setTitle("Add Item");
        }
    }

    private void initViews() {
        etName = findViewById(R.id.etItemName);
        etQuantity = findViewById(R.id.etItemQuantity);
        etPrice = findViewById(R.id.etItemPrice);
        etLowStockThreshold = findViewById(R.id.etLowStockThreshold);
        spinnerCategory = findViewById(R.id.spinnerCategory);
        ivItemImage = findViewById(R.id.ivItemImage);
        btnCamera = findViewById(R.id.btnCamera);
        btnGallery = findViewById(R.id.btnGallery);
        btnSave = findViewById(R.id.btnSave);

        btnCamera.setOnClickListener(v -> checkCameraPermission());
        btnGallery.setOnClickListener(v -> checkStoragePermission());
        btnSave.setOnClickListener(v -> saveItem());
    }

    private void setupCategorySpinner() {
        ArrayAdapter<String> categoryAdapter = new ArrayAdapter<>(
                this, android.R.layout.simple_spinner_item, Constants.ITEM_CATEGORIES);
        categoryAdapter.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item);
        spinnerCategory.setAdapter(categoryAdapter);
    }

    private void loadItem(int itemId) {
        executor.execute(() -> {
            currentItem = itemDao.getItemById(itemId);
            if (currentItem != null) {
                runOnUiThread(() -> {
                    etName.setText(currentItem.getName());
                    etQuantity.setText(String.valueOf(currentItem.getQuantity()));
                    etPrice.setText(String.valueOf(currentItem.getPrice()));
                    etLowStockThreshold.setText(String.valueOf(currentItem.getLowStockThreshold()));

                    // Set category spinner
                    for (int i = 0; i < Constants.ITEM_CATEGORIES.length; i++) {
                        if (Constants.ITEM_CATEGORIES[i].equals(currentItem.getCategory())) {
                            spinnerCategory.setSelection(i);
                            break;
                        }
                    }

                    // Load image
                    currentImagePath = currentItem.getImagePath();
                    if (currentImagePath != null && new File(currentImagePath).exists()) {
                        Glide.with(this)
                                .load(new File(currentImagePath))
                                .centerCrop()
                                .into(ivItemImage);
                    }
                });
            }
        });
    }

    private void checkCameraPermission() {
        if (ContextCompat.checkSelfPermission(this, Manifest.permission.CAMERA)
                != PackageManager.PERMISSION_GRANTED) {
            ActivityCompat.requestPermissions(this,
                    new String[]{Manifest.permission.CAMERA}, PERMISSION_CAMERA);
        } else {
            openCamera();
        }
    }

    private void checkStoragePermission() {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU) {
            if (ContextCompat.checkSelfPermission(this, Manifest.permission.READ_MEDIA_IMAGES)
                    != PackageManager.PERMISSION_GRANTED) {
                ActivityCompat.requestPermissions(this,
                        new String[]{Manifest.permission.READ_MEDIA_IMAGES}, PERMISSION_STORAGE);
            } else {
                openGallery();
            }
        } else {
            if (ContextCompat.checkSelfPermission(this, Manifest.permission.READ_EXTERNAL_STORAGE)
                    != PackageManager.PERMISSION_GRANTED) {
                ActivityCompat.requestPermissions(this,
                        new String[]{Manifest.permission.READ_EXTERNAL_STORAGE}, PERMISSION_STORAGE);
            } else {
                openGallery();
            }
        }
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions,
                                           @NonNull int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);
        if (grantResults.length > 0 && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
            if (requestCode == PERMISSION_CAMERA) {
                openCamera();
            } else if (requestCode == PERMISSION_STORAGE) {
                openGallery();
            }
        } else {
            Toast.makeText(this, "Permission denied", Toast.LENGTH_SHORT).show();
        }
    }

    private void openCamera() {
        Intent takePictureIntent = new Intent(MediaStore.ACTION_IMAGE_CAPTURE);
        if (takePictureIntent.resolveActivity(getPackageManager()) != null) {
            File photoFile = null;
            try {
                photoFile = ImageUtils.createImageFile(this);
                cameraPhotoPath = photoFile.getAbsolutePath();
            } catch (IOException e) {
                Toast.makeText(this, "Error creating image file", Toast.LENGTH_SHORT).show();
                return;
            }

            Uri photoUri = FileProvider.getUriForFile(this,
                    getApplicationContext().getPackageName() + ".fileprovider", photoFile);
            takePictureIntent.putExtra(MediaStore.EXTRA_OUTPUT, photoUri);
            startActivityForResult(takePictureIntent, REQUEST_CAMERA);
        }
    }

    private void openGallery() {
        Intent intent = new Intent(Intent.ACTION_PICK, MediaStore.Images.Media.EXTERNAL_CONTENT_URI);
        startActivityForResult(intent, REQUEST_GALLERY);
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);

        if (resultCode == RESULT_OK) {
            if (requestCode == REQUEST_CAMERA && cameraPhotoPath != null) {
                currentImagePath = cameraPhotoPath;
                Glide.with(this)
                        .load(new File(currentImagePath))
                        .centerCrop()
                        .into(ivItemImage);
            } else if (requestCode == REQUEST_GALLERY && data != null && data.getData() != null) {
                Uri selectedImageUri = data.getData();
                String savedPath = ImageUtils.saveImageFromUri(this, selectedImageUri);
                if (savedPath != null) {
                    currentImagePath = savedPath;
                    Glide.with(this)
                            .load(new File(currentImagePath))
                            .centerCrop()
                            .into(ivItemImage);
                } else {
                    Toast.makeText(this, "Failed to save image", Toast.LENGTH_SHORT).show();
                }
            }
        }
    }

    private void saveItem() {
        String name = etName.getText().toString().trim();
        String quantityStr = etQuantity.getText().toString().trim();
        String priceStr = etPrice.getText().toString().trim();
        String thresholdStr = etLowStockThreshold.getText().toString().trim();
        String category = spinnerCategory.getSelectedItem().toString();

        // Validation
        if (name.isEmpty()) {
            etName.setError("Name is required");
            etName.requestFocus();
            return;
        }
        if (quantityStr.isEmpty()) {
            etQuantity.setError("Quantity is required");
            etQuantity.requestFocus();
            return;
        }
        if (priceStr.isEmpty()) {
            etPrice.setError("Price is required");
            etPrice.requestFocus();
            return;
        }

        int quantity;
        double price;
        int threshold;
        try {
            quantity = Integer.parseInt(quantityStr);
            price = Double.parseDouble(priceStr);
            threshold = thresholdStr.isEmpty() ? Constants.DEFAULT_LOW_STOCK_THRESHOLD
                    : Integer.parseInt(thresholdStr);
        } catch (NumberFormatException e) {
            Toast.makeText(this, "Please enter valid numbers", Toast.LENGTH_SHORT).show();
            return;
        }

        if (isEditMode && currentItem != null) {
            currentItem.setName(name);
            currentItem.setCategory(category);
            currentItem.setQuantity(quantity);
            currentItem.setPrice(price);
            currentItem.setLowStockThreshold(threshold);
            currentItem.setImagePath(currentImagePath);

            executor.execute(() -> {
                itemDao.update(currentItem);
                runOnUiThread(() -> {
                    Toast.makeText(this, "Item updated", Toast.LENGTH_SHORT).show();
                    finish();
                });
            });
        } else {
            Item newItem = new Item();
            newItem.setName(name);
            newItem.setCategory(category);
            newItem.setQuantity(quantity);
            newItem.setPrice(price);
            newItem.setLowStockThreshold(threshold);
            newItem.setImagePath(currentImagePath);

            executor.execute(() -> {
                itemDao.insert(newItem);
                runOnUiThread(() -> {
                    Toast.makeText(this, "Item added", Toast.LENGTH_SHORT).show();
                    finish();
                });
            });
        }
    }

    @Override
    public boolean onSupportNavigateUp() {
        finish();
        return true;
    }

    @Override
    protected void onDestroy() {
        super.onDestroy();
        if (executor != null && !executor.isShutdown()) {
            executor.shutdown();
        }
    }
}
