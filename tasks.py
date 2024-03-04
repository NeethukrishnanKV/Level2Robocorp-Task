from robocorp.tasks import task
from robocorp import browser
import Utilities
@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    Utilities.open_robot_order_website()
    orders = Utilities.get_orders()
    Utilities.fill_the_form(orders)
    Utilities.archive_receipts()
