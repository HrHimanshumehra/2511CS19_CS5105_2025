# Inventory Manager - Android App

A feature-rich inventory management Android application built with Java, XML, and Room Database.

## Features

### 1. Custom App Icon
- Warehouse/box themed adaptive icon (API 26+)
- PNG fallback icons for all screen densities (mdpi through xxxhdpi)
- Green color scheme matching the app theme

### 2. Item Image Support
- Take photos using the device camera
- Pick images from the gallery
- Images stored locally with automatic resizing and compression
- Thumbnail display in inventory list cards

### 3. Partial Stock Removal
- Dedicated "Use Stock" button on each item card
- Enter units to remove (e.g., sold 3 units, consumed 5 units)
- Validates against current stock level
- Automatically triggers low stock warning dialog when quantity drops below threshold

### 4. Sort & Filter by Category
- Chip-based category filter bar: All, Electronics, Food, Tools, Clothing, Office Supplies, Other
- Sort options: Date Added, Name, Quantity, Price
- Combined filter + sort for fast navigation of large inventories

## Tech Stack

- **Language:** Java
- **UI:** XML layouts with Material Design Components
- **Database:** Room (SQLite)
- **Image Loading:** Glide
- **Min SDK:** 24 (Android 7.0)
- **Target SDK:** 34 (Android 14)

## Project Structure

```
app/src/main/java/com/inventory/manager/
├── MainActivity.java              # Dashboard with list, filter, sort
├── AddEditItemActivity.java       # Add/edit item with image capture
├── adapter/
│   └── ItemAdapter.java           # RecyclerView adapter with action buttons
├── database/
│   ├── AppDatabase.java           # Room database singleton
│   └── ItemDao.java               # Data access object with sort/filter queries
├── model/
│   └── Item.java                  # Room entity with fields and low stock check
└── utils/
    └── ImageUtils.java            # Image save, resize, delete utilities
```

## How to Build

1. Open the `InventoryManager` folder in Android Studio
2. Sync Gradle files
3. Run on an emulator or physical device (API 24+)
