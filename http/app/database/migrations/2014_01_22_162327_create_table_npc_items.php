<?php

use Illuminate\Database\Migrations\Migration;

class CreateTableNpcItems extends Migration {
  public function up() {
    Schema::create('npc_items', function($table) {
      $table->increments('id');
      $table->integer('npc_id')->unsigned();
      $table->integer('item_id')->unsigned();
      
      $table->integer('value')->unsigned();
      $table->boolean('chance');
      
      $table->foreign('npc_id')
             ->references('id')
             ->on('npcs');
      
      $table->foreign('item_id')
             ->references('id')
             ->on('items');
    });
  }
  
  public function down() {
    Schema::drop('npc_items');
  }
}