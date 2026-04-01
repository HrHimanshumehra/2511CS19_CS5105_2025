package com.inventory.manager.database;

import androidx.room.Dao;
import androidx.room.Delete;
import androidx.room.Insert;
import androidx.room.Query;
import androidx.room.Update;

import com.inventory.manager.model.Item;

import java.util.List;

@Dao
public interface ItemDao {

    @Insert
    long insert(Item item);

    @Update
    void update(Item item);

    @Delete
    void delete(Item item);

    @Query("SELECT * FROM items ORDER BY dateAdded DESC")
    List<Item> getAllItems();

    @Query("SELECT * FROM items WHERE id = :id")
    Item getItemById(int id);

    @Query("SELECT * FROM items WHERE category = :category ORDER BY dateAdded DESC")
    List<Item> getItemsByCategory(String category);

    // Sort options
    @Query("SELECT * FROM items ORDER BY name ASC")
    List<Item> getAllSortedByName();

    @Query("SELECT * FROM items ORDER BY quantity ASC")
    List<Item> getAllSortedByQuantity();

    @Query("SELECT * FROM items ORDER BY price DESC")
    List<Item> getAllSortedByPrice();

    @Query("SELECT * FROM items ORDER BY dateAdded DESC")
    List<Item> getAllSortedByDate();

    // Sort + Filter by category
    @Query("SELECT * FROM items WHERE category = :category ORDER BY name ASC")
    List<Item> getByCategorySortedByName(String category);

    @Query("SELECT * FROM items WHERE category = :category ORDER BY quantity ASC")
    List<Item> getByCategorySortedByQuantity(String category);

    @Query("SELECT * FROM items WHERE category = :category ORDER BY price DESC")
    List<Item> getByCategorySortedByPrice(String category);

    @Query("SELECT * FROM items WHERE category = :category ORDER BY dateAdded DESC")
    List<Item> getByCategorySortedByDate(String category);

    @Query("SELECT DISTINCT category FROM items ORDER BY category ASC")
    List<String> getAllCategories();
}
