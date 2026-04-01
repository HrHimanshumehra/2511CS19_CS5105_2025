package com.inventory.manager.adapter;

import android.content.Context;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.Button;
import android.widget.ImageButton;
import android.widget.ImageView;
import android.widget.TextView;

import androidx.annotation.NonNull;
import androidx.cardview.widget.CardView;
import androidx.core.content.ContextCompat;
import androidx.recyclerview.widget.RecyclerView;

import com.bumptech.glide.Glide;
import com.inventory.manager.R;
import com.inventory.manager.model.Item;

import java.io.File;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

public class ItemAdapter extends RecyclerView.Adapter<ItemAdapter.ItemViewHolder> {

    private List<Item> items = new ArrayList<>();
    private final Context context;
    private final OnItemActionListener listener;

    public interface OnItemActionListener {
        void onEditItem(Item item);
        void onDeleteItem(Item item);
        void onUseStock(Item item);
    }

    public ItemAdapter(Context context, OnItemActionListener listener) {
        this.context = context;
        this.listener = listener;
    }

    public void setItems(List<Item> items) {
        this.items = items;
        notifyDataSetChanged();
    }

    @NonNull
    @Override
    public ItemViewHolder onCreateViewHolder(@NonNull ViewGroup parent, int viewType) {
        View view = LayoutInflater.from(parent.getContext())
                .inflate(R.layout.item_card, parent, false);
        return new ItemViewHolder(view);
    }

    @Override
    public void onBindViewHolder(@NonNull ItemViewHolder holder, int position) {
        Item item = items.get(position);

        holder.tvName.setText(item.getName());
        holder.tvCategory.setText(item.getCategory());
        holder.tvQuantity.setText(String.format(Locale.getDefault(), "Qty: %d", item.getQuantity()));

        NumberFormat currencyFormat = NumberFormat.getCurrencyInstance(Locale.getDefault());
        holder.tvPrice.setText(currencyFormat.format(item.getPrice()));

        // Load image
        if (item.getImagePath() != null && new File(item.getImagePath()).exists()) {
            Glide.with(context)
                    .load(new File(item.getImagePath()))
                    .centerCrop()
                    .placeholder(R.drawable.ic_placeholder)
                    .into(holder.ivItemImage);
        } else {
            holder.ivItemImage.setImageResource(R.drawable.ic_placeholder);
        }

        // Low stock warning
        if (item.isLowStock()) {
            holder.tvLowStock.setVisibility(View.VISIBLE);
            holder.cardView.setCardBackgroundColor(
                    ContextCompat.getColor(context, R.color.low_stock_background));
        } else {
            holder.tvLowStock.setVisibility(View.GONE);
            holder.cardView.setCardBackgroundColor(
                    ContextCompat.getColor(context, R.color.card_background));
        }

        // Button listeners
        holder.btnUseStock.setOnClickListener(v -> listener.onUseStock(item));
        holder.btnEdit.setOnClickListener(v -> listener.onEditItem(item));
        holder.btnDelete.setOnClickListener(v -> listener.onDeleteItem(item));
    }

    @Override
    public int getItemCount() {
        return items.size();
    }

    static class ItemViewHolder extends RecyclerView.ViewHolder {
        CardView cardView;
        ImageView ivItemImage;
        TextView tvName, tvCategory, tvQuantity, tvPrice, tvLowStock;
        Button btnUseStock;
        ImageButton btnEdit, btnDelete;

        ItemViewHolder(@NonNull View itemView) {
            super(itemView);
            cardView = itemView.findViewById(R.id.cardView);
            ivItemImage = itemView.findViewById(R.id.ivItemImage);
            tvName = itemView.findViewById(R.id.tvItemName);
            tvCategory = itemView.findViewById(R.id.tvItemCategory);
            tvQuantity = itemView.findViewById(R.id.tvItemQuantity);
            tvPrice = itemView.findViewById(R.id.tvItemPrice);
            tvLowStock = itemView.findViewById(R.id.tvLowStock);
            btnUseStock = itemView.findViewById(R.id.btnUseStock);
            btnEdit = itemView.findViewById(R.id.btnEdit);
            btnDelete = itemView.findViewById(R.id.btnDelete);
        }
    }
}
