import uuid
import related
from sqlalchemy.dialects.postgresql import JSONB

from sqlalchemy import (
    BigInteger,
    Boolean,
    Column,
    DateTime,
    ForeignKey,
    Integer,
    String,
    Text,
    UniqueConstraint,
    distinct,
    dialects,
    func,
)

from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship
from sqlalchemy_utils import JSONType, ScalarListType, UUIDType

Base = declarative_base()

class Job(Base):
    __tablename__ = "job"

    id = Column(BigInteger().with_variant(Integer, "sqlite"), primary_key=True)
    user = Column(Text, nullable=False)  # must be present in cognito, only owner can edit the job
    job_id = Column(
        UUIDType(binary=False),
        unique=True,
        default=lambda: str(uuid.uuid4().hex),
        index=True,
    )
    submission_deadline = Column(
        DateTime(timezone=True), server_default=func.now(), index=True
    )
    meta_data = Column(JSONType, server_default="{}")
    tags = Column(ScalarListType, server_default="[]")
    date_created = Column(DateTime(timezone=True), server_default=func.now(), index=True)
    message_priority = Column(
        Integer, default=3
    )  # expected values 1=high 2=medium 3=low
    last_modified_date = Column(DateTime(timezone=True), server_default=func.now(), index=True)

    # relationship
    status_id = Column(Integer, ForeignKey("status.id"), index=True)

class File(Base):
    __tablename__ = "file"

    __table_args__ = (UniqueConstraint("sha1", "job_id"),)

    id = Column(BigInteger().with_variant(Integer, "sqlite"), primary_key=True)
    sha1 = Column(String(40), index=True, nullable=False)
    md5 = Column(String(32))
    last_modified_date = Column(DateTime(timezone=True), server_default=func.now())
    source = Column(Text)
    meta_data = Column(JSONB, server_default="{}")
    date_created = Column(DateTime(timezone=True), server_default=func.now())
    s3_location = Column(Text)
    status = Column(Text, server_default="UNKNOWN")

    # duplicate from job due to join limitation w falcon-autocrud
    user = Column(
        Text, nullable=False
    )  # must be present in cognito, only owner can edit the job
    job_id = Column(UUIDType(binary=False), nullable=False)  # ensure job_id is valid

class Status(Base):
    __tablename__ = "status"

    id = Column(Integer, primary_key=True)
    label = Column(Text, nullable=False)
    jobs = relationship("Job", backref="status")

    def __repr__(self):
        return f"<Status(id={self.id}, label='{self.label}')>"

@related.mutable()
class UserDetails(object):
    username = related.StringField(required=False)
    sla = related.IntegerField(required=False)
    date_created = related.StringField(required=False)
    status = related.StringField(required=False)
    role = related.StringField(required=False)
