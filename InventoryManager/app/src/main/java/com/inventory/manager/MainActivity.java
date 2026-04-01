package com.inventory.manager;

import android.content.Intent;
import android.os.Bundle;
import android.view.View;
import android.widget.AdapterView;
import android.widget.ArrayAdapter;
import android.widget.EditText;
import android.widget.Spinner;
import android.widget.TextView;
import android.widget.Toast;

import androidx.appcompat.app.AlertDialog;
import androidx.appcompat.app.AppCompatActivity;
import androidx.recyclerview.widget.LinearLayoutManager;
import androidx.recyclerview.widget.RecyclerView;

import com.google.android.material.chip.Chip;
import com.google.android.material.chip.ChipGroup;
import com.google.android.material.floatingactionbutton.FloatingActionButton;
import com.inventory.manager.adapter.ItemAdapter;
import com.inventory.manager.database.AppDatabase;
import com.inventory.manager.database.ItemDao;
import com.inventory.manager.model.Item;
import com.inventory.manager.utils.Constants;
import com.inventory.manager.utils.ImageUtils;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class MainActivity extends AppCompatActivity implements ItemAdapter.OnItemActionListener {

    private static final int REQUEST_ADD_ITEM = 1;
    private static final int REQUEST_EDIT_ITEM = 2;

    private RecyclerView recyclerView;
    private ItemAdapter adapter;
    private ItemDao itemDao;
    private ExecutorService executor;
    private TextView tvEmptyState;
    private ChipGroup chipGroupCategories;
    private Spinner spinnerSort;

    private String selectedCategory = Constants.CATEGORY_ALL;
    private int selectedSortIndex = 0; // 0=Date, 1=Name, 2=Quantity, 3=Price

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        setSupportActionBar(findViewById(R.id.toolbar));

        executor = Executors.newSingleThreadExecutor();
        itemDao = AppDatabase.getInstance(this).itemDao();

        initViews();
        setupCategoryChips();
        setupSortSpinner();
    }

    @Override
    protected void onResume() {
        super.onResume();
        loadItems();
    }

    private void initViews() {
        recyclerView = findViewById(R.id.recyclerView);
        tvEmptyState = findViewById(R.id.tvEmptyState);
        chipGroupCategories = findViewById(R.id.chipGroupCategories);
        spinnerSort = findViewById(R.id.spinnerSort);

        recyclerView.setLayoutManager(new LinearLayoutManager(this));
        adapter = new ItemAdapter(this, this);
        recyclerView.setAdapter(adapter);

        FloatingActionButton fab = findViewById(R.id.fabAddItem);
        fab.setOnClickListener(v -> {
            Intent intent = new Intent(MainActivity.this, AddEditItemActivity.class);
            startActivityForResult(intent, REQUEST_ADD_ITEM);
        });
    }

    private void setupCategoryChips() {
        for (String category : Constants.FILTER_CATEGORIES) {
            Chip chip = new Chip(this);
            chip.setText(category);
            chip.setCheckable(true);
            chip.setCheckedIconVisible(true);

            if (category.equals("All")) {
                chip.setChecked(true);
            }

            chipGroupCategories.addView(chip);
        }

        chipGroupCategories.setSingleSelection(true);
        chipGroupCategories.setSelectionRequired(true);

        chipGroupCategories.setOnCheckedStateChangeListener((group, checkedIds) -> {
            if (!checkedIds.isEmpty()) {
                Chip selectedChip = group.findViewById(checkedIds.get(0));
                if (selectedChip != null) {
                    selectedCategory = selectedChip.getText().toString();
                    loadItems();
                }
            }
        });
    }

    private void setupSortSpinner() {
        ArrayAdapter<String> sortAdapter = new ArrayAdapter<>(
                this, android.R.layout.simple_spinner_item, Constants.SORT_OPTIONS);
        sortAdapter.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item);
        spinnerSort.setAdapter(sortAdapter);

        spinnerSort.setOnItemSelectedListener(new AdapterView.OnItemSelectedListener() {
            @Override
            public void onItemSelected(AdapterView<?> parent, View view, int position, long id) {
                selectedSortIndex = position;
                loadItems();
            }

            @Override
            public void onNothingSelected(AdapterView<?> parent) {
                // No action needed
            }
        });
    }

    private void loadItems() {
        executor.execute(() -> {
            List<Item> items;

            if (selectedCategory.equals(Constants.CATEGORY_ALL)) {
                switch (selectedSortIndex) {
                    case 1:
                        items = itemDao.getAllSortedByName();
                        break;
                    case 2:
                        items = itemDao.getAllSortedByQuantity();
                        break;
                    case 3:
                        items = itemDao.getAllSortedByPrice();
                        break;
                    default:
                        items = itemDao.getAllSortedByDate();
                        break;
                }
            } else {
                switch (selectedSortIndex) {
                    case 1:
                        items = itemDao.getByCategorySortedByName(selectedCategory);
                        break;
                    case 2:
                        items = itemDao.getByCategorySortedByQuantity(selectedCategory);
                        break;
                    case 3:
                        items = itemDao.getByCategorySortedByPrice(selectedCategory);
                        break;
                    default:
                        items = itemDao.getByCategorySortedByDate(selectedCategory);
                        break;
                }
            }

            List<Item> finalItems = items != null ? items : new ArrayList<>();

            runOnUiThread(() -> {
                adapter.setItems(finalItems);
                tvEmptyState.setVisibility(finalItems.isEmpty() ? View.VISIBLE : View.GONE);
                recyclerView.setVisibility(finalItems.isEmpty() ? View.GONE : View.VISIBLE);
            });
        });
    }

    @Override
    public void onEditItem(Item item) {
        Intent intent = new Intent(this, AddEditItemActivity.class);
        intent.putExtra("item_id", item.getId());
        startActivityForResult(intent, REQUEST_EDIT_ITEM);
    }

    @Override
    public void onDeleteItem(Item item) {
        new AlertDialog.Builder(this)
                .setTitle("Delete Item")
                .setMessage("Are you sure you want to delete \"" + item.getName() + "\"?")
                .setPositiveButton("Delete", (dialog, which) -> {
                    executor.execute(() -> {
                        ImageUtils.deleteImage(item.getImagePath());
                        itemDao.delete(item);
                        runOnUiThread(this::loadItems);
                    });
                    Toast.makeText(this, "Item deleted", Toast.LENGTH_SHORT).show();
                })
                .setNegativeButton("Cancel", null)
                .show();
    }

    @Override
    public void onUseStock(Item item) {
        View dialogView = getLayoutInflater().inflate(R.layout.dialog_use_stock, null);
        EditText etAmount = dialogView.findViewById(R.id.etUseAmount);
        TextView tvCurrentStock = dialogView.findViewById(R.id.tvCurrentStock);
        TextView tvItemNameDialog = dialogView.findViewById(R.id.tvItemNameDialog);

        tvItemNameDialog.setText(item.getName());
        tvCurrentStock.setText(String.format("Current stock: %d units", item.getQuantity()));

        new AlertDialog.Builder(this)
                .setTitle("Use Stock")
                .setView(dialogView)
                .setPositiveButton("Remove", (dialog, which) -> {
                    String amountStr = etAmount.getText().toString().trim();
                    if (amountStr.isEmpty()) {
                        Toast.makeText(this, "Please enter an amount", Toast.LENGTH_SHORT).show();
                        return;
                    }

                    int amount;
                    try {
                        amount = Integer.parseInt(amountStr);
                    } catch (NumberFormatException e) {
                        Toast.makeText(this, "Please enter a valid number", Toast.LENGTH_SHORT).show();
                        return;
                    }
                    if (amount <= 0) {
                        Toast.makeText(this, "Amount must be greater than 0", Toast.LENGTH_SHORT).show();
                        return;
                    }

                    if (amount > item.getQuantity()) {
                        Toast.makeText(this, "Cannot remove more than current stock", Toast.LENGTH_SHORT).show();
                        return;
                    }

                    item.setQuantity(item.getQuantity() - amount);

                    executor.execute(() -> {
                        itemDao.update(item);
                        runOnUiThread(() -> {
                            loadItems();
                            if (item.isLowStock()) {
                                new AlertDialog.Builder(this)
                                        .setTitle("⚠ Low Stock Warning")
                                        .setMessage("\"" + item.getName() + "\" is now at " +
                                                item.getQuantity() + " units, which is at or below " +
                                                "the threshold of " + item.getLowStockThreshold() + " units.")
                                        .setPositiveButton("OK", null)
                                        .show();
                            }
                            Toast.makeText(this,
                                    amount + " units removed. Remaining: " + item.getQuantity(),
                                    Toast.LENGTH_SHORT).show();
                        });
                    });
                })
                .setNegativeButton("Cancel", null)
                .show();
    }

    @Override
    protected void onDestroy() {
        super.onDestroy();
        if (executor != null && !executor.isShutdown()) {
            executor.shutdown();
        }
    }
}
