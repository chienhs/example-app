<?php

use App\Http\Controllers\aaaa;
use Illuminate\Support\Facades\Route;

Route::get('/', function () {
    return view('welcome');
});

Route::get('/test', [aaaa::class, 'newexportTemplateKeNgang'])->name('aaaa');
