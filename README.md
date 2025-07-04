# Introduction
This Script is to collect relevant social media posts from xiaohongshu.com<br>
input: keywords，output: two spreadsheets, search_result.xlsx，notes_contents and comments.xlsx<br>
Targeted info: username, ip (if available), posydate, post content, post comments replied by the author<br>
Recommend to contact the XiaoHongShu to get permission by email RED.AD@xiaohongshu.com <br>

# Instruction
Run the script and follow the prompts<br>

# Update log
20250405:<br>#
fix bug: indefinete loop origining from the change of the loading end prompt in XiaoHongShu website; other small bugs. <br>
20250205:<br>
optimise the updates checking mechanism <br>
20241206:<br>
lower the data collection frequency to avoid any burden to XiaoHongShu website <be>
enhance the fault-tolerant mechanism: ignore not targeted posts and save them in notes_save_wrong.xlsx <be>
add the stop-and-pick-up function for data collection <br>
20241130：<br>
publish.