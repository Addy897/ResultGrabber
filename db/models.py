from sqlalchemy import Column, String, Integer, Float, ForeignKey, CheckConstraint, PrimaryKeyConstraint, CHAR
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship

Base = declarative_base()

class Student(Base):
    __tablename__ = 'Student'
    USN = Column(String(20), primary_key=True)
    StudentName = Column(String(50), nullable=False)

class Subject(Base):
    __tablename__ = 'Subject'
    SubjectCode = Column(String(10), primary_key=True)
    SubjectName = Column(String(50), nullable=False)
    Semester = Column(Integer, nullable=False)

class Result(Base):
    __tablename__ = 'Result'
    USN = Column(String(20), ForeignKey('Student.USN'))
    SubjectCode = Column(String(10), ForeignKey('Subject.SubjectCode'))
    SEE = Column(Integer)
    CIE = Column(Integer)
    Total = Column(Integer)
    Result = Column(CHAR(1))

    __table_args__ = (
        PrimaryKeyConstraint('USN', 'SubjectCode'),
        CheckConstraint('SEE BETWEEN 0 AND 50'),
        CheckConstraint('CIE BETWEEN 0 AND 50'),
        CheckConstraint("Result IN ('P', 'F')")
    )
