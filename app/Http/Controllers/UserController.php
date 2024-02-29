<?php

namespace App\Http\Controllers;

use Illuminate\Foundation\Auth\Access\AuthorizesRequests;
use Illuminate\Foundation\Bus\DispatchesJobs;
use Illuminate\Foundation\Validation\ValidatesRequests;
use Illuminate\Http\Request;
use Illuminate\Routing\Controller as BaseController;
use Illuminate\Support\Env;
use Illuminate\Support\Facades\Config;

class UserController extends Controller {
    use AuthorizesRequests, DispatchesJobs, ValidatesRequests;

    public function balance(Request $request) {
    }
}
