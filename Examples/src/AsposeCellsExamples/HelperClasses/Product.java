package AsposeCellsExamples.HelperClasses;
// ExStart: 1
public class Product
{
	String name;
	int quantity;

	public Product(String name, int quantity)
	{
		this.quantity = quantity;
		this.name = name;
	}

	public int getQuantity()
	{
		return this.quantity;
	}

	public void setQuantity(int value)
	{
		this.quantity = value;
	}

	public String getName()
	{
		return this.name;
	}

	public void setName(String value)
	{
		this.name = value;
	}
}
// ExEnd: 1