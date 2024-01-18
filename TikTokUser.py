from TikTokApi import TikTokApi
import asyncio
import os
import xlsxwriter

workbook = xlsxwriter.Workbook('Creator_Info.xlsx')
worksheet = workbook.add_worksheet()

ms_token = os.environ.get(
    "ms_token", None
) 
context_options = {
    'viewport' : { 'width': 1280, 'height': 1024 },
    'user_agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.150 Safari/537.36'
}

async def trending_videos():
    async with TikTokApi() as api:
        await api.create_sessions(ms_tokens=[ms_token], num_sessions=1, sleep_after=3, context_options=context_options)

        # Change the tag name for different type of videos
        tag = api.hashtag(name="toy")
        row = 0
        async for video in tag.videos(count=1000):
            # print(video)
            # print(video.as_dict)
            user_name = video.author.username
            print(user_name)
            user = api.user(user_name)
            user_data = await user.info()
            user_info = user_data["userInfo"]
            follower_count = user_info["stats"]["followerCount"]
            signature = user_info["user"]["signature"]
            worksheet.write(row, 0, user_name)
            worksheet.write(row, 1, follower_count)
            worksheet.write(row, 2, signature)
            print(follower_count)
            print("Signature: " + signature)
            print()
            row += 1

    workbook.close()

if __name__ == "__main__":
    asyncio.run(trending_videos())