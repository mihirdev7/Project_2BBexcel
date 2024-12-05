package com.example.excel3

import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.TextView
import androidx.recyclerview.widget.RecyclerView

class CustomerAdapter(
    private val customers: List<Customer>
) : RecyclerView.Adapter<CustomerAdapter.CustomerViewHolder>() {

    inner class CustomerViewHolder(itemView: View) : RecyclerView.ViewHolder(itemView) {
        val customerName: TextView = itemView.findViewById(R.id.customerNameTextView)
        val customerEmail: TextView = itemView.findViewById(R.id.customerEmailTextView)
        val customerPhone: TextView = itemView.findViewById(R.id.customerPhoneTextView)
        val customerAddress: TextView = itemView.findViewById(R.id.customerAddressTextView)
    }

    override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): CustomerViewHolder {
        val itemView = LayoutInflater.from(parent.context)
            .inflate(R.layout.item_customer, parent, false)
        return CustomerViewHolder(itemView)
    }

    override fun onBindViewHolder(holder: CustomerViewHolder, position: Int) {
        val customer = customers[position]
        holder.customerName.text = customer.name
        holder.customerEmail.text = customer.email
        holder.customerPhone.text = customer.phone
        holder.customerAddress.text = customer.address
    }

    override fun getItemCount(): Int {
        return customers.size
    }
}
