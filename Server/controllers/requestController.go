package controllers

import (
	"net/http"
	"project-its/initializers"
	"project-its/models"

	"github.com/gin-gonic/gin"
)

func RequestIndex(c *gin.Context) {

	var request []models.BookingRapat
	// Tambahkan filter untuk tidak menampilkan event dengan status "pending"
	if err := initializers.DB.Where("status = ?", "pending").Find(&request).Error; err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}
	c.JSON(http.StatusOK, gin.H{"request": request})

}

func AccRequest(c *gin.Context) {

	id := c.Params.ByName("id")

	var request models.BookingRapat

	if err := initializers.DB.First(&request, id); err.Error != nil {
		c.JSON(404, gin.H{"error": "Request not found"})
		return
	}

	// Cek apakah statusnya "pending"
	if request.Status != "pending" {
		c.JSON(400, gin.H{"error": "Request status is not pending"})
		return
	}

	initializers.DB.Model(&request).Update("status", "acc")

	c.JSON(200, gin.H{"message": "Request accepted"})

}

func CancelRequest(c *gin.Context) {

	id := c.Params.ByName("id")

	var request models.BookingRapat

	if err := initializers.DB.First(&request, id); err.Error != nil {
		c.JSON(404, gin.H{"error": "Request not found"})
		return
	}

	if err := initializers.DB.Delete(&request).Error; err != nil {
		c.JSON(400, gin.H{"error": "Failed to delete Request: " + err.Error()})
		return
	}

	c.Status(204)

}
