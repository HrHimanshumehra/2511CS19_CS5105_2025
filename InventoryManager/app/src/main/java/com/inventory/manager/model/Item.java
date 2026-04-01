package com.inventory.manager.model;

import androidx.room.Entity;
import androidx.room.PrimaryKey;

@Entity(tableName = "items")
public class Item {

    @PrimaryKey(autoGenerate = true)
    private int id;

    private String name;
    private String category;
    private int quantity;
    private double price;
    private String imagePath;
    private int lowStockThreshold;
    private long dateAdded;

    public Item() {
        this.dateAdded = System.currentTimeMillis();
        this.lowStockThreshold = 5;
    }

    // Getters
    public int getId() { return id; }
    public String getName() { return name; }
    public String getCategory() { return category; }
    public int getQuantity() { return quantity; }
    public double getPrice() { return price; }
    public String getImagePath() { return imagePath; }
    public int getLowStockThreshold() { return lowStockThreshold; }
    public long getDateAdded() { return dateAdded; }

    // Setters
    public void setId(int id) { this.id = id; }
    public void setName(String name) { this.name = name; }
    public void setCategory(String category) { this.category = category; }
    public void setQuantity(int quantity) { this.quantity = quantity; }
    public void setPrice(double price) { this.price = price; }
    public void setImagePath(String imagePath) { this.imagePath = imagePath; }
    public void setLowStockThreshold(int lowStockThreshold) { this.lowStockThreshold = lowStockThreshold; }
    public void setDateAdded(long dateAdded) { this.dateAdded = dateAdded; }

    public boolean isLowStock() {
        return quantity <= lowStockThreshold;
    }
}
