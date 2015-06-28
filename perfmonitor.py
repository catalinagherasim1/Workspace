from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
from sqlalchemy import create_engine, Column, Integer, String, update, ForeignKey
from sqlalchemy import MetaData
import wmi
import time
import win32com.client

Base = declarative_base()

class Computer(Base):
    __tablename__ = "Computer"
    id = Column(Integer, primary_key=True)
    comp_name = Column(String(50),nullable=False)

    networks = relationship("Network", backref="Computer")
    disks = relationship("Disk", backref="Computer")
    processors = relationship("Processor", backref="Computer")
    memory = relationship("Memory", backref="Computer")

class Memory(Base):
    __tablename__ = "Memory"
    id = Column(Integer, primary_key=True)
    free_memory = Column(Integer)
    limit_memory = Column(Integer)
    comp_id = Column(Integer, ForeignKey('Computer.id'))


class Network(Base):
	__tablename__ = "Network"
	id = Column(Integer, primary_key=True)
	down_speed = Column(Integer)
	up_speed = Column(Integer)
	bytes_total = Column(Integer)
	net_name = Column(String(150))
	comp_id = Column(Integer, ForeignKey('Computer.id'))


class Disk(Base):
	__tablename__ = "Disk"
	id = Column(Integer, primary_key=True)
	disk_name = Column(String(50))
	disk_time = Column(Integer)
	idle_time = Column(Integer)
	disk_reads = Column(Integer)
	comp_id = Column(Integer, ForeignKey('Computer.id'))


class Processor(Base):
	__tablename__ = "Processor"
	id = Column(Integer, primary_key=True)
	proc_name = Column(String(50))
	proc_time = Column(Integer)
	comp_id = Column(Integer, ForeignKey('Computer.id'))

engine = create_engine('mysql+mysqlconnector://root:@localhost/statistics', echo=False)

Base.metadata.drop_all(bind=engine)
Base.metadata.create_all(engine)

Session = sessionmaker(bind=engine)
session = Session()

disklist = []
proclist = []
netlist = []
memlist = []
objlist = []

f = open('computers.txt', 'r')
for line in f.readlines():
    cname = line.strip()
    conectat_cu_succes = 1
    print 'Incerc sa ma conectez la', cname
    try:
        WMIService = wmi.WMI(cname)
    except:
        print "Nu am reusit sa ma conectez la", cname
        conectat_cu_succes = 0

    if conectat_cu_succes==1:
        print 'M-am conectat cu succes la', cname
        obj = win32com.client.Dispatch("WbemScripting.SWbemRefresher")
        objlist.append(obj)
        diskItems = obj.AddEnum(WMIService, "Win32_PerfFormattedData_PerfDisk_PhysicalDisk").objectSet
        disklist.append(diskItems)
        processorItems = obj.AddEnum(WMIService, "Win32_PerfFormattedData_PerfOS_Processor").objectSet
        proclist.append(processorItems)
        networkItems = obj.AddEnum(WMIService, "Win32_PerfFormattedData_Tcpip_NetworkInterface").objectSet
        netlist.append(networkItems)
        memoryItems = obj.AddEnum(WMIService, "Win32_PerfFormattedData_PerfOS_Memory").objectSet
        memlist.append(memoryItems)
        computer1 = Computer(id=len(disklist), comp_name=cname)
        session.add(computer1)
        session.commit()

if len(disklist)==0:
    print "Nu am reusit sa ma conectez la niciun calculator"
    exit()

while 1:

    disk_list = session.query(Disk).all()
    for elem in disk_list:
        session.delete(elem)

    proc_list = session.query(Processor).all()
    for elem in proc_list:
        session.delete(elem)

    net_list = session.query(Network).all()
    for elem in net_list:
        session.delete(elem)

    mem_list = session.query(Memory).all()
    for elem in mem_list:
        session.delete(elem)

    session.commit()
    
    for pos in range(0,len(disklist)):
        print "Computer #", pos+1 
        objlist[pos].Refresh()

        print "Disks:"
        for item in disklist[pos]:
            disk1 = Disk(disk_name=item.Name, disk_time=item.PercentDiskTime, idle_time=item.PercentIdleTime, disk_reads=item.DiskReadsPerSec, comp_id=pos+1)
            session.add(disk1)
            print item.Name, " ",
            print item.PercentDiskTime,
            print item.PercentIdleTime,
            print item.DiskReadsPerSec

        print "Processors:"
        for item in proclist[pos]:
            processor1 = Processor(proc_name=item.Name, proc_time=item.PercentProcessorTime, comp_id=pos+1)
            session.add(processor1)
            print item.Name, " ",
            print item.PercentProcessorTime, "%"

        print "Memory:"
        for item in memlist[pos]:
            memory1 = Memory(free_memory=item.AvailableMBytes, limit_memory=int(item.CommitLimit)/(1024**2), comp_id=pos+1)
            session.add(memory1)
            print item.AvailableMBytes,
            print int(item.CommitLimit)/(1024**2)

        print "Networks:"
        for item in netlist[pos]:
            network1 = Network(net_name=item.Name, down_speed=item.BytesReceivedPerSec, up_speed=item.BytesSentPerSec, bytes_total=item.BytesTotalPerSec, comp_id=pos+1)

            session.add(network1)
            print item.Name,
            print "Down", item.BytesReceivedPerSec, "B/s",
            print "Up", item.BytesSentPerSec, "B/s",
            print "Total", item.BytesTotalPerSec, "B/s"
        session.commit()

    time.sleep(5)
