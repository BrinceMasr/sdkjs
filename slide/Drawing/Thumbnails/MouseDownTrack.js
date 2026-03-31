function CMouseDownTrack(Thumbnails)
{
	this.Thumbnails = Thumbnails;
	this.SetDefault();
}
CMouseDownTrack.prototype.Check = function()
{
	if(this.Started)
	{
		let aThPages = this.Thumbnails.m_arrPages;
		if(!aThPages[this.Page])
		{
			this.Reset();
		}
	}
};
CMouseDownTrack.prototype.SetDefault = function()
{
	this.Started = false;
	this.StartedSimple = true;
	this.Page = -1;
	this.X = -1;
	this.Y = -1;
	this.Position = -1;
};
CMouseDownTrack.prototype.Reset = function()
{
	this.SetDefault();
};
CMouseDownTrack.prototype.Start = function(Page, X, Y)
{
	this.Started = true;
	this.StartedSimple = true;
	this.Page = Page;
	this.X = X;
	this.Y = Y;
};
CMouseDownTrack.prototype.IsStarted = function()
{
	this.Check();
	return this.Started;
};
CMouseDownTrack.prototype.IsSimple = function()
{
	this.Check();
	return this.StartedSimple;
};
CMouseDownTrack.prototype.IsDragged = function()
{
	return this.IsStarted() && !this.IsSimple() && this.GetPosition() !== -1;
};
CMouseDownTrack.prototype.ResetSimple = function(Position)
{
	this.StartedSimple = false;
	this.Position = Position;
};
CMouseDownTrack.prototype.SetPosition = function(Position)
{
	this.Position = Position;
};
CMouseDownTrack.prototype.GetPosition = function()
{
	this.Check();
	return this.Position;
};
CMouseDownTrack.prototype.GetPage = function()
{
	this.Check();
	return this.Page;
};
CMouseDownTrack.prototype.GetX = function()
{
	this.Check();
	return this.X;
};
CMouseDownTrack.prototype.GetY = function()
{
	this.Check();
	return this.Y;
};

CMouseDownTrack.prototype.IsMoved = function(X, Y)
{
	this.Check();
	if (Math.abs(this.X - X) > 10 || Math.abs(this.Y - Y) > 10)
		return true;
	return false;
};
CMouseDownTrack.prototype.IsSamePos = function()
{
	this.Check();
	return this.Position === this.Page || this.Position === (this.Page + 1);
};

function COutlineMouseDownTrack(Thumbnails)
{
	this.Thumbnails = Thumbnails;
	this.SetDefault();
}
COutlineMouseDownTrack.prototype.Check = function()
{
	if(this.Started)
	{
		let aThPages = this.Thumbnails.m_arrPages;
		if(!aThPages[this.Page])
		{
			this.Reset();
		}
	}
};
COutlineMouseDownTrack.prototype.SetDefault = function()
{
	this.Started = false;
	this.StartedSimple = true;
	this.Page = -1;
	this.X = -1;
	this.Y = -1;
	this.Position = 1;
};
COutlineMouseDownTrack.prototype.Reset = function()
{
	this.SetDefault();
};
COutlineMouseDownTrack.prototype.Start = function(Page, X, Y)
{
	this.Started = true;
	this.StartedSimple = true;
	this.Page = Page;
	this.X = X;
	this.Y = Y;
};
COutlineMouseDownTrack.prototype.IsStarted = function()
{
	this.Check();
	return this.Started;
};
COutlineMouseDownTrack.prototype.IsSimple = function()
{
	this.Check();
	return this.StartedSimple;
};
COutlineMouseDownTrack.prototype.IsDragged = function()
{
	return this.IsStarted() && !this.IsSimple() && this.GetPosition() !== -1;
};
COutlineMouseDownTrack.prototype.ResetSimple = function(Position)
{
	this.StartedSimple = false;
	this.Position = Position;
};
COutlineMouseDownTrack.prototype.SetPosition = function(Position)
{
	this.Position = Position;
};
COutlineMouseDownTrack.prototype.GetPosition = function()
{
	this.Check();
	return this.Position;
};
COutlineMouseDownTrack.prototype.GetPage = function()
{
	this.Check();
	return this.Page;
};
COutlineMouseDownTrack.prototype.GetX = function()
{
	this.Check();
	return this.X;
};
COutlineMouseDownTrack.prototype.GetY = function()
{
	this.Check();
	return this.Y;
};

COutlineMouseDownTrack.prototype.IsMoved = function(X, Y)
{
	this.Check();
	if (Math.abs(this.X - X) > 10 || Math.abs(this.Y - Y) > 10)
		return true;
	return false;
};
COutlineMouseDownTrack.prototype.IsSamePos = function()
{
	this.Check();
	return this.Position === this.Page || this.Position === (this.Page + 1);
};

window["AscCommon"] = window["AscCommon"] || {};
window["AscCommon"].CMouseDownTrack = CMouseDownTrack;
window["AscCommon"].COutlineMouseDownTrack = COutlineMouseDownTrack;