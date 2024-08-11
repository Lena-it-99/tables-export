<?php

use App\Models\Order;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Route;

Route::apiResource('orders', OrderController::class);

Route::get('/orders', function () {
    return Order::all();
});

Route::get('/orders/{id}', function ($id) {
    return Order::find($id);
});

Route::post('/orders', function (Request $request) {
    return Order::create($request->all());
});

Route::put('/orders/{id}', function (Request $request, $id) {
    $order = Order::findOrFail($id);
    $order->update($request->all());

    return $order;
});

Route::delete('/orders/{id}', function ($id) {
    Order::find($id)->delete();

    return 204;
});