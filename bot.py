import discord
from discord.ext import commands
from discord import app_commands
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import os
import re
import traceback
from dotenv import load_dotenv

load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ── Channel → User mapping for Woolix auto-read ───────────────────────────────
WOOLIX_CHANNELS = {
    1425584567292792995: "Slater",
    1451859096499257356: "Nuke",
}

def get_sheet():
    creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(os.getenv("SHEET_ID"))

def get_log_sheet():
    return get_sheet().worksheet("Order Log")

def get_summary_sheet():
    return get_sheet().worksheet("Summary")

def get_orders(sheet):
    """Returns list of (row_index, order_num, row_data) for all logged orders"""
    all_rows = sheet.get_all_values()
    orders = []
    order_num = 0
    for i, row in enumerate(all_rows[3:], start=4):  # skip 3 header rows
        placed_by = row[3] if len(row) > 3 else ""  # column D
        if placed_by in ("Slater", "Nuke"):
            order_num += 1
            orders.append((i, order_num, row))
    return orders

def next_empty_row(sheet):
    col = sheet.col_values(4)  # column D (Placed By)
    for i, val in enumerate(col[3:], start=4):
        if not val:
            return i
    return len(col) + 1

def parse_total(raw: str) -> float:
    """Strip $ and commas, return float. Returns 0.0 on failure."""
    try:
        return float(re.sub(r"[^\d.]", "", raw))
    except:
        return 0.0

# ── Bot Setup ─────────────────────────────────────────────────────────────────
intents = discord.Intents.default()
intents.message_content = True   # required to read embed content in on_message
bot = commands.Bot(command_prefix="!", intents=intents)
tree = bot.tree

@bot.event
async def on_ready():
    await tree.sync()
    print(f"✅ {bot.user} is online and synced.")

# ── Woolix Embed Auto-Reader ──────────────────────────────────────────────────
@bot.event
async def on_message(message: discord.Message):
    # Only watch the two designated channels
    if message.channel.id not in WOOLIX_CHANNELS:
        await bot.process_commands(message)
        return

    # Must have embeds
    if not message.embeds:
        await bot.process_commands(message)
        return

    for embed in message.embeds:
        title = embed.title or ""

        # Detect Woolix checkout success embed
        if "Order Successfully Placed" not in title:
            continue

        placed_by = WOOLIX_CHANNELS[message.channel.id]

        # Build a field lookup dict  {field_name: field_value}
        fields = {f.name: f.value for f in embed.fields}

        store           = fields.get("Store", "Unknown Store")
        total_raw       = fields.get("Total", "0")
        order_items     = fields.get("Order Items", "")
        eta             = fields.get("Estimated Arrival", "")
        customer_name   = fields.get("Name", "")
        phone           = fields.get("Phone", "")
        delivery_addr   = fields.get("Delivery Address", "")
        session_id      = fields.get("Woolix Session ID", "")

        order_total = parse_total(total_raw)

        # Notes column: store + items summary
        notes = f"{store} | {order_items[:80]}{'...' if len(order_items) > 80 else ''}"

        try:
            sheet = get_log_sheet()
            row = next_empty_row(sheet)
            orders = get_orders(sheet)
            order_num = len(orders) + 1
            now = datetime.now()
            date = now.strftime("%m/%d/%Y")
            time_str = now.strftime("%I:%M %p")

            # Sheet columns: B=Date, C=Time, D=Placed By, E=Order Total, F=Pre-Auth, G=Settled, H=Credit Added, I=Notes
            sheet.update([[
                date,
                time_str,
                placed_by,
                order_total,
                order_total,  # pre-auth = total (UberEats holds full amount)
                "",           # settled — filled later via /settle
                "",           # credit added
                notes
            ]], f"B{row}:I{row}")

            # Confirmation embed
            color = 0x00B894 if placed_by == "Slater" else 0xFDCB6E
            confirm = discord.Embed(
                title=f"📋 Auto-Logged — Order #{order_num}",
                color=color
            )
            confirm.add_field(name="Placed By",  value=placed_by,          inline=True)
            confirm.add_field(name="Store",       value=store,              inline=True)
            confirm.add_field(name="Total",       value=f"${order_total:.2f}", inline=True)
            if customer_name:
                confirm.add_field(name="Customer", value=customer_name,    inline=True)
            if eta:
                confirm.add_field(name="ETA",      value=eta,              inline=True)
            if delivery_addr:
                confirm.add_field(name="Address",  value=delivery_addr,    inline=True)
            confirm.add_field(name="Status", value="⏳ Pending delivery", inline=False)
            confirm.set_footer(text=f"Order #{order_num} • Auto-logged from Woolix • {date} {time_str}")

            await message.channel.send(embed=confirm)

        except Exception:
            await message.channel.send(
                f"⚠️ Woolix order detected but failed to log:\n```{traceback.format_exc()}```"
            )

    await bot.process_commands(message)

# ── /order ────────────────────────────────────────────────────────────────────
@tree.command(name="order", description="Log a new order")
@app_commands.describe(
    placed_by="Who placed the order (Slater or Nuke)",
    order_total="Total order amount in dollars",
    preauth_hold="Pre-auth hold amount placed by the platform"
)
async def log_order(
    interaction: discord.Interaction,
    placed_by: str,
    order_total: float,
    preauth_hold: float
):
    await interaction.response.defer()
    try:
        sheet = get_log_sheet()
        row = next_empty_row(sheet)
        orders = get_orders(sheet)
        order_num = len(orders) + 1
        now = datetime.now()
        date = now.strftime("%m/%d/%Y")
        time_str = now.strftime("%I:%M %p")

        sheet.update([[
            date,
            time_str,
            placed_by.capitalize(),
            order_total,
            preauth_hold,
            "",
            "",
            ""
        ]], f"B{row}:I{row}")

        embed = discord.Embed(
            title="📋 Order Logged",
            color=0x00B894 if placed_by.lower() == "slater" else 0xFDCB6E
        )
        embed.add_field(name="Order #",      value=str(order_num),        inline=True)
        embed.add_field(name="Placed By",    value=placed_by.capitalize(), inline=True)
        embed.add_field(name="Order Total",  value=f"${order_total:.2f}", inline=True)
        embed.add_field(name="Pre-Auth Hold",value=f"${preauth_hold:.2f}",inline=True)
        embed.add_field(name="Status",       value="⏳ Pending delivery", inline=True)
        embed.set_footer(text=f"Order #{order_num} • Row {row} • {date} {time_str}")
        await interaction.followup.send(embed=embed)

    except Exception:
        await interaction.followup.send(f"❌ Error logging order:\n```{traceback.format_exc()}```")

# ── /settle ───────────────────────────────────────────────────────────────────
@tree.command(name="settle", description="Mark an order as delivered and log the settled charge")
@app_commands.describe(
    order_number="The order number to settle",
    settled_amount="The actual charge after delivery"
)
async def settle_order(
    interaction: discord.Interaction,
    order_number: int,
    settled_amount: float
):
    await interaction.response.defer()
    try:
        sheet = get_log_sheet()
        orders = get_orders(sheet)

        target_row = None
        for (row_idx, order_num, row_data) in orders:
            if order_num == order_number:
                target_row = row_idx
                break

        if not target_row:
            await interaction.followup.send(f"❌ Order #{order_number} not found.")
            return

        sheet.update_cell(target_row, 7, settled_amount)  # column G

        embed = discord.Embed(title="✅ Order Settled", color=0x00B894)
        embed.add_field(name="Order #",        value=str(order_number),      inline=True)
        embed.add_field(name="Settled Charge", value=f"${settled_amount:.2f}", inline=True)
        embed.add_field(name="Status",         value="✅ Delivered",          inline=True)
        await interaction.followup.send(embed=embed)

    except Exception:
        await interaction.followup.send(f"❌ Error settling order:\n```{traceback.format_exc()}```")

# ── /addcredit ────────────────────────────────────────────────────────────────
@tree.command(name="addcredit", description="Log credit you added to the card")
@app_commands.describe(amount="Amount of credit added in dollars")
async def add_credit(interaction: discord.Interaction, amount: float):
    await interaction.response.defer()
    try:
        sheet = get_log_sheet()
        row = next_empty_row(sheet)
        now = datetime.now()

        sheet.update([[
            now.strftime("%m/%d/%Y"),
            now.strftime("%I:%M %p"),
            "Slater",
            "",
            "",
            "",
            amount,
            "Credit top-up"
        ]], f"B{row}:I{row}")

        embed = discord.Embed(title="💳 Credit Added", color=0x6C5CE7)
        embed.add_field(name="Amount Added", value=f"${amount:.2f}", inline=True)
        embed.add_field(name="Added By",     value="Slater",         inline=True)
        embed.set_footer(text=now.strftime("%m/%d/%Y %I:%M %p"))
        await interaction.followup.send(embed=embed)

    except Exception:
        await interaction.followup.send(f"❌ Error logging credit:\n```{traceback.format_exc()}```")

# ── /balance ──────────────────────────────────────────────────────────────────
@tree.command(name="balance", description="Check current balance and spend breakdown")
async def check_balance(interaction: discord.Interaction):
    await interaction.response.defer()
    try:
        summary = get_summary_sheet()

        def get_val(cell):
            try:
                v = summary.acell(cell).value
                return float(str(v).replace("$", "").replace(",", "")) if v else 0.0
            except:
                return 0.0

        capital_one   = get_val("B17")
        active_holds  = get_val("B19")
        true_usable   = get_val("B20")
        slater_credit = get_val("B7")
        nuke_credit   = get_val("B8")
        total_settled = get_val("B14")
        slater_orders = summary.acell("B12").value or "0"
        nuke_orders   = summary.acell("B13").value or "0"

        embed = discord.Embed(title="📊 Slater Services — Balance Overview", color=0x0F3460)
        embed.add_field(
            name="💳 Capital One",
            value=f"Available: **${capital_one:.2f}**\nActive Holds: **${active_holds:.2f}**\n✅ True Usable: **${true_usable:.2f}**",
            inline=False
        )
        embed.add_field(
            name="🟢 Slater",
            value=f"Orders: **{slater_orders}**\nCredit Added: **${slater_credit:.2f}**",
            inline=True
        )
        embed.add_field(
            name="🟡 Nuke",
            value=f"Orders: **{nuke_orders}**\nCredit Added: **${nuke_credit:.2f}**",
            inline=True
        )
        embed.add_field(name="📦 Total Settled Charges", value=f"**${total_settled:.2f}**", inline=False)
        embed.set_footer(text=f"Updated {datetime.now().strftime('%m/%d/%Y %I:%M %p')}")
        await interaction.followup.send(embed=embed)

    except Exception:
        await interaction.followup.send(f"❌ Error fetching balance:\n```{traceback.format_exc()}```")

# ── /updatebalance ────────────────────────────────────────────────────────────
@tree.command(name="updatebalance", description="Update the Capital One available credit shown in the app")
@app_commands.describe(amount="Current available credit showing on Capital One")
async def update_balance(interaction: discord.Interaction, amount: float):
    await interaction.response.defer()
    try:
        summary = get_summary_sheet()
        summary.update_acell("B17", amount)

        embed = discord.Embed(title="💳 Capital One Balance Updated", color=0xFDCB6E)
        embed.add_field(name="Available Credit", value=f"**${amount:.2f}**", inline=True)
        embed.set_footer(text=datetime.now().strftime("%m/%d/%Y %I:%M %p"))
        await interaction.followup.send(embed=embed)

    except Exception:
        await interaction.followup.send(f"❌ Error updating balance:\n```{traceback.format_exc()}```")

# ── Run ───────────────────────────────────────────────────────────────────────
bot.run(os.getenv("DISCORD_TOKEN"))
