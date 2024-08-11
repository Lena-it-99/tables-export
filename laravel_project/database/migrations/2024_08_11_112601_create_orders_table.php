<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

class CreateOrdersTable extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('orders', function (Blueprint $table) {
            $table->id(); // المفتاح الأساسي
            $table->string('customer');
            $table->string('item');
            $table->string('material');
            $table->integer('quantity');
            $table->date('date');
            $table->string('barcode');
            $table->string('customer_phone');
            $table->float('back_height');
            $table->float('box_height');
            $table->float('legs_height');
            $table->float('mattress_height');
            $table->boolean('additional_cups');
            $table->string('colors');
            $table->text('additional_notes')->nullable();
            $table->string('user');
            $table->float('total_price');
            $table->float('remaining');
            $table->date('proposed_delivery_date');
            $table->string('city');
            $table->string('region');
            $table->string('city_neighborhood');
            $table->string('order_status');
            $table->string('address');
            $table->string('piece_id');
            $table->string('source');
            $table->string('delivery_type');
            $table->timestamps();
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('orders');
    }
}