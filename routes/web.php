<?php

use Illuminate\Support\Facades\Route;
use \App\Http\Controllers\UserController;
use \App\Http\Controllers\PptController;
use \App\Http\Controllers\PowerPointController;
/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

Route::get('/', function () {
    return view('welcome');
});

Route::middleware([])->group(function () {
    Route::get('/', function () {
        // 使用 first 和 second 中间件
    });

    Route::get('/balance/{id?}', [UserController::class, 'balance']);
    Route::get('/index', [PptController::class, 'index']);
    Route::get('/xy', [PptController::class, 'xy']);
    Route::get('/t', [PptController::class, 't']);
    Route::get('/txt', [PptController::class, 'txt']);
    Route::get('/line', [PptController::class, 'line']);
    Route::get('/ppt', [PowerPointController::class, 'ppt']);
    Route::get('/two', [PowerPointController::class, 'two']);
    Route::get('/four', [PowerPointController::class, 'four']);
    Route::get('/five', [PowerPointController::class, 'five']);
    Route::get('/six', [PowerPointController::class, 'six']);
    Route::get('/seven', [PowerPointController::class, 'seven']);
    Route::get('/eight', [PowerPointController::class, 'eight']);
    Route::get('/nine', [PowerPointController::class, 'nine']);

});
