package com.inventory.manager.utils;

public final class Constants {

    private Constants() {}

    public static final String[] ITEM_CATEGORIES = {
            "Electronics", "Food", "Tools", "Clothing", "Office Supplies", "Other"
    };

    public static final String[] FILTER_CATEGORIES = {
            "All", "Electronics", "Food", "Tools", "Clothing", "Office Supplies", "Other"
    };

    public static final String[] SORT_OPTIONS = {
            "Date Added", "Name", "Quantity", "Price"
    };

    public static final String CATEGORY_ALL = "All";
    public static final int DEFAULT_LOW_STOCK_THRESHOLD = 5;
}
