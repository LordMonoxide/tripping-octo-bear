<?php

use Illuminate\Database\Migrations\Migration;

class CreateTableUserItems extends Migration {
  public function up() {
    Schema::create('user_items', function($table) {
      $table->increments('id');
      $table->integer('item_id')->unsigned();
      $table->integer('value')->unsigned();
      $table->boolean('bound');
      $table->timestamps();
      
      $table->foreign('item_id')
             ->references('id')
             ->on('items');
    });
  }
  
  public function down() {
    Schema::drop('user_items');
  }
}