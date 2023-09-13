export class Fruit {
  name: string;
  family: string;
  nutritions: {
    calories: number;
    fat: number;
    sugar: number;
    carbohydrates: number;
    protein: number;
  };

  constructor(
    name: string,
    family: string,
    calories: number,
    fat: number,
    sugar: number,
    carbohydrates: number,
    protein: number
  ) {
    this.name = name;
    this.family = family;
    this.nutritions.calories = calories;
    this.nutritions.fat = fat;
    this.nutritions.sugar = sugar;
    this.nutritions.carbohydrates = carbohydrates;
    this.nutritions.protein = protein
  }
}
