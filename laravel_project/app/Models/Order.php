<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Order extends Model
{
    use HasFactory;

    protected $fillable = [
        'customer', 
        'item', 
        'material', 
        'quantity', 
        'date', 
        'barcode', 
        'customer_phone', 
        'back_height', 
        'box_height', 
        'legs_height', 
        'mattress_height', 
        'additional_cups', 
        'colors', 
        'additional_notes', 
        'user', 
        'total_price', 
        'remaining', 
        'proposed_delivery_date', 
        'city', 
        'region', 
        'city_neighborhood', 
        'order_status', 
        'address', 
        'piece_id', 
        'source', 
        'delivery_type'
    ];
}
