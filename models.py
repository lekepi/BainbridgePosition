from sqlalchemy import create_engine, ForeignKey, Column, Integer, String, Date, Float, Boolean
from sqlalchemy.orm import relationship, sessionmaker
from sqlalchemy.ext.declarative import declarative_base
from config import ConfigDefault
from datetime import datetime
import sys

config_class = ConfigDefault  # get the config file from config.py classes

try:
    my_URI = config_class.SQLALCHEMY_DATABASE_URI
except:
    print("The config file is incorrect")
    sys.exit(1)

Base = declarative_base()
engine = create_engine(my_URI)
Session = sessionmaker()
Session.configure(bind=engine)
session = Session()


class Trade(Base):
    __tablename__ = 'trade'
    id = Column(Integer, primary_key=True)
    order_number = Column(String(length=20), nullable=False)
    trade_date = Column(Date, nullable=False)
    settle_date = Column(Date, nullable=False)
    side = Column(String(length=1), nullable=False)
    is_short = Column(Boolean)
    quantity = Column(Float)
    ticker = Column(String(length=20), nullable=False)
    exec_price = Column(Float, nullable=False)
    broker = Column(String(length=10))
    account = Column(String(length=45))
    cncy = Column(String(length=3))
    sec_name = Column(String(length=45))
    sedol = Column(String(length=7))
    isin = Column(String(length=12))
    cusip = Column(String(length=20))
    bbg_type = Column(String(length=10))
    is_cfd = Column(Boolean)
    origin = Column(String(length=10))
    active = Column(Boolean, default=True)
    created_time = Column(Date, nullable=False, default=datetime.utcnow)
    modified_time = Column(Date)
    created_by = Column(String(length=20), nullable=False)
    modified_by = Column(String(length=20))
    comment = Column(String(length=200))
    file_name = Column(String(length=200))
    product_id = Column(ForeignKey("product.id"))
    product = relationship("Product")
    parent_fund_id = Column(ForeignKey("parent_fund.id"))
    parent_fund = relationship("ParentFund")
    long_future_name = Column(String(length=45))
    allocs = relationship('Allocation', viewonly=True, lazy=True)
    parent_broker_id = Column(ForeignKey("parent_broker.id"))
    parent_broker = relationship("ParentBroker")

    def __repr__(self):
        return f"Trade: {self.trade_date.strftime('%Y-%m-%d')}, id:{self.id}, Account:{self.account} -" \
               f" {self.side} {int(self.quantity)} {self.ticker}"


class Allocation(Base):
    __tablename__ = 'allocation'
    id = Column(Integer, primary_key=True)
    trade_id = Column(ForeignKey("trade.id"))
    trade = relationship("Trade")
    account_id = Column(ForeignKey("account.id"))
    account = relationship("Account")
    quantity = Column(Float)
    active = Column(Boolean, default=True)
    exec_fee = Column(Float)
    sec_fee = Column(Float)


class ParentBroker(Base):
    __tablename__ = 'parent_broker'
    id = Column(Integer, primary_key=True)
    name = Column(String(length=10), nullable=False)
    long_name = Column(String(length=45), nullable=False)


class OmsBroker(Base):
    __tablename__ = 'oms_broker'
    id = Column(Integer, primary_key=True)
    code = Column(String(length=20), nullable=False)
    parent_broker_id = Column(ForeignKey("parent_broker.id"))
    parent_broker = relationship("ParentBroker")


class FeeRule(Base):
    __tablename__ = 'fee_rule'
    id = Column(Integer, primary_key=True)
    name = Column(String(length=45), nullable=False)
    rate = Column(Float, nullable=False)
    start_date = Column(Date, nullable=False)


class ExecFee(Base):
    __tablename__ = 'exec_fee'
    id = Column(Integer, primary_key=True)
    parent_broker_id = Column(ForeignKey("parent_broker.id"))
    parent_broker = relationship("ParentBroker")
    country_id = Column(ForeignKey("country.id"))
    country = relationship("Country")
    rate = Column(Float, nullable=False)
    rate_type = Column(String(length=10), nullable=False)
    min_amount = Column(Float, nullable=False)
    start_date = Column(Date, nullable=False)


class Exchange(Base):
    __tablename__ = 'exchange'
    id = Column(Integer, primary_key=True)
    bbg_code = Column(String(length=10), nullable=False)
    country_id = Column(ForeignKey("country.id"))
    country = relationship("Country")


class Country(Base):
    __tablename__ = 'country'
    id = Column(Integer, primary_key=True)
    name = Column(String(length=45), nullable=False)
    continent = Column(String(length=10), nullable=False)
    iso_code = Column(String(length=3), nullable=False)
    settle_day = Column(Integer)
    holidays = relationship('SettlementHoliday', viewonly=True, lazy=True)


class SettlementHoliday(Base):
    __tablename__ = 'settlement_holiday'
    id = Column(Integer, primary_key=True)
    country_id = Column(ForeignKey("country.id"), nullable=False)
    country = relationship("Country")
    holiday_date = Column(Date, nullable=False)


class IndustrySector(Base):
    __tablename__ = 'industry_sector'
    id = Column(Integer, primary_key=True)
    name = Column(String(45))


class IndustryGroup(Base):
    __tablename__ = 'industry_group'
    id = Column(Integer, primary_key=True)
    name = Column(String(45))


class Account(Base):
    __tablename__ = 'account'
    id = Column(Integer, primary_key=True)
    asset_class = Column(String(length=45))
    location = Column(String(length=45))
    broker = Column(String(length=45))
    code = Column(String(length=45))
    name = Column(String(length=45))
    fund_id = Column(ForeignKey("fund.id"))
    fund = relationship("Fund")
    oms_account_id = Column(ForeignKey("oms_account.id"))
    oms_account = relationship("OmsAccount")


class OmsAccount(Base):
    __tablename__ = 'oms_account'
    id = Column(Integer, primary_key=True)
    code = Column(String(length=20))
    still_active = Column(Boolean)
    parent_fund_id = Column(ForeignKey("parent_fund.id"))
    parent_fund = relationship("ParentFund")


class ParentFund(Base):
    __tablename__ = 'parent_fund'
    id = Column(Integer, primary_key=True)
    name = Column(String(length=45))


class Fund(Base):
    __tablename__ = 'fund'
    id = Column(Integer, primary_key=True)
    name = Column(String(length=45))
    parent = Column(String(length=45))


class FundSplit(Base):
    __tablename__ = 'fund_split'
    id = Column(Integer, primary_key=True)
    start_date = Column(Date, nullable=False)
    client = Column(String(length=45), nullable=False)
    fund_id = Column(ForeignKey("fund.id"), nullable=False)
    fund = relationship("Fund")
    percentage = Column(Float, nullable=False)


class TaskChecker(Base):
    __tablename__ = 'task_checker'
    id = Column(Integer, primary_key=True)
    date_time = Column(Date, nullable=False, default=datetime.utcnow)
    task_name = Column(String(length=45), nullable=False)
    task_type = Column(String(length=45), nullable=False)
    task_details = Column(String(length=45), nullable=False)
    status = Column(String(length=45), nullable=False)
    comment = Column(String(length=200), nullable=False)
    active = Column(Boolean, default=True)

    def __repr__(self):
        return f"<TaskChecker(date='{self.date_time}', task name='{self.task_name}', status='{self.status}')>"


class LogDb(Base):
    __tablename__ = 'log_db'
    id = Column(Integer, primary_key=True)
    date_time = Column(Date, nullable=False, default=datetime.utcnow)
    project = Column(String(length=45), nullable=False)
    task = Column(String(length=45), nullable=False)
    issue = Column(String(length=45), nullable=False)
    msg_type = Column(String(length=45), nullable=False)
    description = Column(String(length=400), nullable=False)


class Product(Base):
    __tablename__ = 'product'
    id = Column(Integer, primary_key=True)
    ticker = Column(String(length=40), nullable=False)
    name = Column(String(length=60), nullable=False)
    isin = Column(String(length=12))
    sedol = Column(String(length=7))
    expiry = Column(Date)
    expiry2 = Column(Date)
    strike = Column(Float)
    prod_type = Column(String(length=10))
    currency_id = Column(ForeignKey("currency.id"))
    currency = relationship("Currency", foreign_keys=[currency_id])
    security_id = Column(ForeignKey("security.id"))
    security = relationship("Security")
    exchange_id = Column(ForeignKey("exchange.id"))
    exchange = relationship("Exchange")
    industry_sector_id = Column(ForeignKey("industry_sector.id"))
    industry_sector = relationship("IndustrySector")
    industry_group_id = Column(ForeignKey("industry_group.id"))
    industry_group = relationship("IndustryGroup")
    currency2_id = Column(ForeignKey("currency.id"))
    currency2 = relationship("Currency", foreign_keys=[currency2_id])
    is_cent = Column(Boolean)
    multiplier = Column(Float)

    def __repr__(self):
        return f"<Product(ticker='{self.ticker}', name='{self.name}', prod_type='{self.prod_type}')>"


class Security(Base):
    __tablename__ = 'security'
    id = Column(Integer, primary_key=True)
    name = Column(String(length=60), nullable=False)
    ticker = Column(String(length=40))
    isin = Column(String(length=12))
    asset_type = Column(String(length=20))
    generic_future = Column(String(length=20))

    def __repr__(self):
        return f"<Product(name='{self.name}', asset_type='{self.asset_type}')>"


class Currency(Base):
    __tablename__ = 'currency'
    id = Column(Integer, primary_key=True)
    name = Column(String(length=45))
    code = Column(String(length=45), nullable=False, unique=True)
    symbol = Column(String(length=45))

    def __repr__(self):
        return f"<Currency(name='{self.name}', code='{self.code}', symbol='{self.symbol}')>"


class PositionPb(Base):
    __tablename__ = 'position_pb'
    id = Column(Integer, primary_key=True)
    entry_date = Column(Date, nullable=False)
    parent_broker_id = Column(ForeignKey("parent_broker.id"))
    parent_broker = relationship("ParentBroker")
    file_name = Column(String(length=100))
    pb_account = Column(String(length=30))
    account_id = Column(ForeignKey("account.id"))
    account = relationship("Account")
    product_id = Column(ForeignKey("product.id"))
    product = relationship("Product")
    isin = Column(String(length=12))
    sedol = Column(String(length=7))
    bbg_ticker = Column(String(length=20))
    ric_ticker = Column(String(length=20))
    sec_name = Column(String(length=45))
    quantity = Column(Float)
    quantity_sd = Column(Float)
    price = Column(Float)
    notional = Column(Float)
    notional_usd = Column(Float)
    cncy = Column(String(length=3))
    currency_id = Column(ForeignKey("currency.id"))
    currency = relationship("Currency", foreign_keys=[currency_id])
    fx_rate = Column(Float)
    is_cfd = Column(Boolean)


class PositionNav(Base):
    __tablename__ = 'position_nav'
    id = Column(Integer, primary_key=True)
    entry_date = Column(Date, nullable=False)
    parent_broker_id = Column(ForeignKey("parent_broker.id"))
    parent_broker = relationship("ParentBroker")
    file_name = Column(String(length=100))
    pb_account = Column(String(length=30))
    account_id = Column(ForeignKey("account.id"))
    account = relationship("Account")
    product_id = Column(ForeignKey("product.id"))
    product = relationship("Product")
    isin = Column(String(length=12))
    sedol = Column(String(length=7))
    bbg_ticker = Column(String(length=20))
    sec_name = Column(String(length=45))
    qty_nav = Column(Float)
    qty_pb = Column(Float)
    price_nav = Column(Float)
    price_pb = Column(Float)
    notional_nav = Column(Float)
    notional_pb = Column(Float)
    notional_nav_usd = Column(Float)
    notional_pb_usd = Column(Float)
    cncy = Column(String(length=3))
    currency_id = Column(ForeignKey("currency.id"))
    currency = relationship("Currency", foreign_keys=[currency_id])
    sec_type = Column(String(length=20))


class NavClassHedge(Base):
    __tablename__ = 'nav_class_hedge'
    id = Column(Integer, primary_key=True)
    entry_date = Column(Date, nullable=False)
    cncy_id = Column(ForeignKey("currency.id"))
    currency = relationship("Currency", foreign_keys=[cncy_id])
    prod_id = Column(ForeignKey("product.id"))
    product = relationship("Product")
    cash_amount = Column(Float)
    hedge_amount = Column(Float)
    cash_pnl = Column(Float)
    hedge_pnl = Column(Float)


class InvestorActivity(Base):
    __tablename__ = 'investor_activity'
    id = Column(Integer, primary_key=True)
    fund_name = Column(String(length=100))
    investor_num = Column(String(length=45))
    investor_name = Column(String(length=45))
    activity = Column(String(length=45))
    activity_mode = Column(String(length=45))
    effective_date = Column(Date)
    cut_off = Column(String(length=45))
    doc_date = Column(Date)
    amount = Column(Float)
    cncy_id = Column(ForeignKey("currency.id"))
    currency = relationship("Currency", foreign_keys=[cncy_id])
    funded_date = Column(Date)
    classes = Column(String(length=45))
    series = Column(String(length=45))


class ExternalCode(Base):
    __tablename__ = 'external_code'
    id = Column(Integer, primary_key=True)
    table_name = Column(String(length=45))
    object_id = Column(Integer)
    code = Column(String(length=45))
    source = Column(String(length=45))


class PositionIssue(Base):
    __tablename__ = 'position_issue'
    id = Column(Integer, primary_key=True)
    entry_date = Column(Date, nullable=False)
    product_id = Column(ForeignKey("product.id"))
    product = relationship("Product")
    parent_fund_id = Column(ForeignKey("parent_fund.id"))
    parent_fund = relationship("ParentFund")
    qty_sys_gs = Column(Float)
    qty_gs = Column(Float)
    qty_sys_ms = Column(Float)
    qty_ms = Column(Float)
    qty_sys_ubs = Column(Float)
    qty_ubs = Column(Float)


class PositionBainb(Base):
    __tablename__ = 'position_bainb'
    id = Column(Integer, primary_key=True)
    entry_date = Column(Date, nullable=False)
    parent_fund_id = Column(ForeignKey("parent_fund.id"))
    parent_fund = relationship("ParentFund")
    ticker = Column(String(length=20))
    product_id = Column(ForeignKey("product.id"))
    product = relationship("Product")
    quantity = Column(Float)
    notional_usd = Column(Float)
    quantity = Column(Float)
    notional_usd = Column(Float)
    trade_qty = Column(Float)
    trade_notional_usd = Column(Float)


def copy_trade(trade):
    amended_trade = Trade(
        order_number=trade.order_number,
        trade_date=trade.trade_date,
        settle_date=trade.settle_date,
        side=trade.side,
        is_short=trade.is_short,
        quantity=trade.quantity,
        ticker=trade.ticker,
        exec_price=trade.exec_price,
        broker=trade.broker,
        account=trade.account,
        cncy=trade.cncy,
        sec_name=trade.sec_name,
        sedol=trade.sedol,
        isin=trade.isin,
        cusip=trade.cusip,
        bbg_type=trade.bbg_type,
        is_cfd=trade.is_cfd,
        origin=trade.origin,
        created_by=trade.created_by,
        product_id=trade.product_id,
        long_future_name=trade.long_future_name,
        oms_account_id=trade.oms_account_id,
        parent_broker_id=trade.parent_broker_id,
    )
    return amended_trade

